VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form RNov1d_frm 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   19605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   19284.03
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid mFlex 
      Height          =   7245
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19621
      _ExtentX        =   34608
      _ExtentY        =   12779
      _Version        =   327680
      Rows            =   1
      Cols            =   8
      FixedCols       =   0
      BackColorFixed  =   13420716
      FillStyle       =   1
      ScrollBars      =   2
      MergeCells      =   2
   End
   Begin VB.Menu mnuPedirMov 
      Caption         =   "Pedir Movil"
      Visible         =   0   'False
      Begin VB.Menu mnuPMoSub 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu mnuPMoSub 
         Caption         =   "Novedad"
         Index           =   1
      End
      Begin VB.Menu mnuPMoSub 
         Caption         =   "Asignar Código "
         Index           =   2
      End
      Begin VB.Menu mnuPMoSub 
         Caption         =   "Cancelar Demora"
         Index           =   3
      End
      Begin VB.Menu mnuPMoSub 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuPMoSub 
         Caption         =   "Pedir AMBU"
         Index           =   5
      End
      Begin VB.Menu mnuPMoSub 
         Caption         =   "Pedir BOMB"
         Index           =   6
      End
      Begin VB.Menu mnuPMoSub 
         Caption         =   "Pedir POLI"
         Index           =   7
      End
      Begin VB.Menu mnuPMoSub 
         Caption         =   "Pedir GEND"
         Index           =   8
      End
      Begin VB.Menu mnuPMoSub 
         Caption         =   "Asignar Móvil"
         Index           =   9
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "M 01"
            Index           =   0
            Tag             =   "M001"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "M 02"
            Index           =   1
            Tag             =   "M002"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "M 03"
            Index           =   2
            Tag             =   "M003"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "M 04"
            Index           =   3
            Tag             =   "M004"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "M 05"
            Index           =   4
            Tag             =   "M005"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "M 06"
            Index           =   5
            Tag             =   "M006"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "M 07"
            Index           =   6
            Tag             =   "M007"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "M 08"
            Index           =   7
            Tag             =   "M008"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "M 09"
            Index           =   8
            Tag             =   "M009"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "M 10"
            Index           =   9
            Tag             =   "M010"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "M 11"
            Index           =   10
            Tag             =   "M011"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "M 12"
            Index           =   11
            Tag             =   "M012"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "M 15"
            Index           =   12
            Tag             =   "M015"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   ""
            Index           =   13
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   ""
            Index           =   14
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "G 00"
            Index           =   15
            Tag             =   "G000"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "G 82"
            Index           =   16
            Tag             =   "G082"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "G 83"
            Index           =   17
            Tag             =   "G083"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "G 84"
            Index           =   18
            Tag             =   "G084"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "G 86"
            Index           =   19
            Tag             =   "G086"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "G 87"
            Index           =   20
            Tag             =   "G087"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "G 88"
            Index           =   21
            Tag             =   "G088"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "G 89"
            Index           =   22
            Tag             =   "G089"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "G 90"
            Index           =   23
            Tag             =   "G090"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "G AUX"
            Index           =   24
            Tag             =   "G"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "GP 01"
            Index           =   25
            Tag             =   "GP01"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "GP 02"
            Index           =   26
            Tag             =   "GP02"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "GP 03"
            Enabled         =   0   'False
            Index           =   27
            Tag             =   "GP03"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "GP 04"
            Enabled         =   0   'False
            Index           =   28
            Tag             =   "GP04"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "GP 05"
            Enabled         =   0   'False
            Index           =   29
            Tag             =   "GP05"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoAMoSub 
            Caption         =   "Móviles No Disponibles"
            Index           =   30
         End
      End
      Begin VB.Menu mnuPMoATar 
         Caption         =   "Asignar Tarea"
         Begin VB.Menu mnuPMoATaSub 
            Caption         =   "Móviles NO Disponibles"
            Index           =   0
         End
         Begin VB.Menu mnuPMoATaSub 
            Caption         =   "M 01"
            Index           =   1
            Tag             =   "M001"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoATaSub 
            Caption         =   "M 02"
            Index           =   2
            Tag             =   "M002"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoATaSub 
            Caption         =   "M 03"
            Index           =   3
            Tag             =   "M003"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoATaSub 
            Caption         =   "M 04"
            Index           =   4
            Tag             =   "M004"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoATaSub 
            Caption         =   "M 05"
            Index           =   5
            Tag             =   "M005"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoATaSub 
            Caption         =   "M 06"
            Index           =   6
            Tag             =   "M006"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoATaSub 
            Caption         =   "M 07"
            Index           =   7
            Tag             =   "M007"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoATaSub 
            Caption         =   "M 08"
            Index           =   8
            Tag             =   "M008"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoATaSub 
            Caption         =   "M 09"
            Index           =   9
            Tag             =   "M009"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoATaSub 
            Caption         =   "M 10"
            Index           =   10
            Tag             =   "M010"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoATaSub 
            Caption         =   "M 11"
            Index           =   11
            Tag             =   "M011"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoATaSub 
            Caption         =   "M 12"
            Index           =   12
            Tag             =   "M012"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPMoATaSub 
            Caption         =   "M 15"
            Index           =   13
            Tag             =   "M015"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuPMoAVa2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPMoAAcc 
         Caption         =   "Marcar como Accidente"
      End
      Begin VB.Menu mnuPMoFEvent 
         Caption         =   "Finalizar Evento"
      End
      Begin VB.Menu mnuPMoAMB 
         Caption         =   "Pedir GCO AMB1"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "RNov1d_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mRec As New ADODB.Recordset
Dim mCodAlfa As String
Dim mKm As String
Dim mSent As String
Dim mReferencia
Dim mResp As Boolean
Dim mPc As String

Private Sub Form_Load()
   mPc = Mid(MDI.mPCname, 1, Len(MDI.mPCname) - 1)
   Inicio
End Sub

Private Sub mFlex_DblClick()
Dim mObj As New clRNov
Dim mCodigo As String
Dim Texto As String
Dim mKm As Double
Dim mFecha As Date
Dim mI As Integer

   mI = mFlex.Col
   Select Case mI
      Case 5
         If mFlex.Text = "SI" Then
            mFlex.Col = 0
            mCodigo = Trim(mFlex.Text)
            mFlex.Col = 1
            mFecha = mFlex.Text
            'Set mRec = mObj.oMovilesCodigo(mCodigo)
            Set mRec = mObj.oMovilesCodigo(mCodigo, mFlex.TextMatrix(mFlex.Row, 1))
            Texto = "Móviles Asignados:   " & NVL(mRec!Mov1, "")
            If mRec!Mov2 <> "" Then
               Texto = Texto & " - " & mRec!Mov2
            End If
            If mRec!Mov3 <> "" Then
               Texto = Texto & " - " & mRec!Mov3
            End If
            mRec.Close
            MsgBox Texto, vbInformation, sMessage
            If mFlex.CellPicture <> 0 Then
               MsgBox "Es una Demora"
            End If
         End If
          
      Case 6 'falta agregar combo ramal
         If mFlex.Row <> 0 Then
            mFlex.Col = 0
            mCodigo = mFlex.Text
            mFlex.Col = 1
            mFecha = mFlex.Text
            Set mRec = mObj.oTabla("novedades2", "WHERE Codigo = '" & mCodigo & "' AND Fecha = '" & Format(mFecha, "yyyy-mm-dd hh:mm:ss") & "' AND CodNov in ('A','B','G','H','M','N','MN','MM','NN')")
            If Not mRec.EOF Then
               RNov6_frm.Show
               For mI = 0 To 1
                  mFlex.Col = mI
                  RNov6_frm.Label4(mI).Caption = mFlex.Text
               Next
               mFlex.Col = 2
               RNov6_frm.Text2.Text = mFlex.Text 'km
               mFlex.Col = 6
               RNov6_frm.Text1.Text = mFlex.Text 'texto novedad
               mFlex.Col = 4 'ramal
               For mI = 0 To RNov6_frm.Combo1(0).ListCount - 1
                  RNov6_frm.Combo1(0).ListIndex = mI
                  If Left(Right((RNov6_frm.Combo1(0).Text), 6), 2) = mFlex.Text Then mI = 99
               
               Next
               RNov6_frm.sLlenoSentido
               mFlex.Col = 3 'ramal
               For mI = 0 To RNov6_frm.Combo1(1).ListCount - 1
                  RNov6_frm.Combo1(1).ListIndex = mI
                  If Left(RNov6_frm.Combo1(1).Text, 2) = mFlex.Text Then mI = 99
               Next
               mFlex.Col = 6
               RNov6_frm.Text1.Text = mFlex.Text
            End If
            mRec.Close
            mFlex.Col = 6
         End If
   End Select
   Set mObj = Nothing
End Sub

Private Sub MFlex_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim mObj As New clRNov
Dim mI As Integer
Dim mJ As Integer
Dim mFlag As Boolean
Dim mCodigo As String
Dim mFecha As String
 
   mFlex.SetFocus
   If Button = 2 Then
      MFlex_MouseDown 1, Shift, X, Y
      If mFlex.Row <> 0 Then
         mFlag = True
         Select Case mFlex.Col
            Case 0
               If mFlex.Text <> "" Then
                  mnuPMoSub(0).Caption = "---- " & mFlex.Text & " ----"
                  mnuPMoSub(0).Tag = ""
                  mnuPMoSub(1).Visible = True
                  mnuPMoSub(2).Visible = False  'Asignar Código
                  mnuPMoSub(3).Visible = False  'Cancelar Demora
                  Set mRec = mObj.oTabla("novedades2", "WHERE Codigo = '" & mFlex.Text & "' AND CodNov in ('MN','AC','CL','FR','FO','SU')")
                  If Not mRec.EOF Then
                     mFlag = False
                  End If
                  mRec.Close
               Else
                  mnuPMoSub(1).Visible = False
                  mnuPMoSub(2).Visible = True
                  For mI = 5 To 8
                     mnuPMoSub(mI).Visible = False
                  Next
               End If
               mnuPMoAAcc.Visible = False
               mnuPMoFEvent.Visible = False
               If mFlag Then
                  mnuPMoSub(3).Visible = False
                  mFlex.Col = 0
                  If fMovDisponibles(RNov1b_frm.Label3, "p6", "p1", "p3") Then
                     mnuPMoAMoSub(30).Visible = False
                  End If
                  If fMovDisponibles(RNov1b_frm.Label4, "g1", "g4", "g5") Then
                     mnuPMoAMoSub(30).Visible = False
                  End If
                  If fMovDisponibles(RNov1b_frm.Label7, "x1", "x1", "x4") Then
                     mnuPMoAMoSub(30).Visible = False
                  End If
                  mnuPMoAAcc.Visible = True
                  'If mFlex.Text <> "" Then
                     mnuPMoFEvent.Visible = True
                  'End If
                  mFlex.Tag = mFlex.Row
                  PopupMenu mnuPedirMov
               End If
               mnuPMoSub(0).Caption = ""
               
            Case 5
               If mFlex.CellPicture <> 0 Then  'DEMORA
                  sInitMenuDemora False
                  mnuPMoSub(3).Visible = True
                  If fMovDisponibles(RNov1b_frm.Label3, "p6", "p1", "p3") Then
                     mnuPMoAMoSub(30).Visible = False
                  End If
                  If fMovDisponibles(RNov1b_frm.Label4, "g1", "g4", "g5") Then
                     mnuPMoAMoSub(30).Visible = False
                  End If
                  If fMovDisponibles(RNov1b_frm.Label7, "x1", "x1", "x4") Then
                     mnuPMoAMoSub(30).Visible = False
                  End If
                  mFlex.Col = 0
                  mnuPMoSub(0).Caption = "---- " & mFlex.Text & " ----"
                  mFlex.Col = 1
                  mnuPMoSub(0).Tag = mFlex.Text
                  mFlex.Col = 4
                  PopupMenu mnuPedirMov
                  sInitMenuDemora True
               End If
         End Select
      End If
   Else
      If mFlex.Row > 0 Then
         mFlex.Col = mFlex.MouseCol
         mFlex.Row = mFlex.MouseRow
      End If
   End If
   Set mObj = Nothing
End Sub




'*****************************************************************
'******  MENU CONTEXTUAL DEL FLEXGRID  ***************************
'*****************************************************************
Private Sub mnuPMoSub_Click(Index As Integer)
Dim mObj As New clRNov
Dim mClima As String
   Select Case Index
      Case 1 'Novedad
          RNov1c_frm.Show
          RNov1c_frm.lCodAlfa.Visible = True
          RNov1c_frm.lCodAlfa.Caption = mFlex.Text
          'mFlex.Col = 2      'mp20160523
          'mKm = mFlex.Text   'mp20160523
          'mFlex.Col = 7      'mp20160523
          'mReferencia = mFlex.Text 'mp20160523
          mKm = fGetKm(mFlex.Text)
          mReferencia = fGetCodigoReferencia(mFlex.Text)
          fPasar_Km_Sent mKm, mFlex.TextMatrix(mFlex.Row, 3), mFlex.TextMatrix(mFlex.Row, 4), mReferencia
          sObtOrigen "", RNov1c_frm.lCodAlfa.Caption, RNov1c_frm.Combo1(1)
          RNov1a_frm.Enabled = False
          RNov1b_frm.Enabled = False
          RNov1d_frm.Enabled = False
      Case 2 'Asignar Código
          mCodAlfa = fNewCodAlfa
          mFlex.Text = mCodAlfa
          mFlex.Col = 1
          mObj.xUpNovedCodAlfa mFlex.Text, mCodAlfa
          Set mObj = Nothing
      Case 3 'Cancelar Demora
          mFlex.Col = 0
          mCodAlfa = mFlex.Text
          mFlex.Col = 1
          Set mRec = mObj.oTabla("novedades2", "where Fecha='" & Format(mFlex.Text, "yyyy-mm-dd hh:mm:ss") & "' and Codigo='" & mCodAlfa & "'")
          If Not mRec.EOF Then
             mClima = ClimaOK(mRec!km)
             mResp = mObj.xInsNovedades("", mCodAlfa, Trim(Right(MDI.mUser, 15)), mRec!km, mRec!sent, "SSV", "CANCELACIÓN DE DEMORA", "T", mClima, "", "", 0, "", "", 0, "", "", 0, "", "", mRec!codramal, mRec!codreferencia) '20160523
             mObj.xUpActualizarNot mPc, 1
             mObj.xUpNovedCodNov mFlex.Text, mCodAlfa, "Q"
             RNov1a_frm.Label4.Visible = mObj.bExistDatoTabla("novedades2", "codnov='D'")
          End If
          mRec.Close
          Set mObj = Nothing
          Unload RNov1d_frm
          Load RNov1d_frm
      Case 5, 6, 7, 8  'Pedido de Móviles Externos
          RNov1c_frm.Show
          RNov1c_frm.Frame1.Caption = "Pedido de Móviles"
          RNov1c_frm.Label2.Caption = Right(mnuPMoSub(Index).Caption, 4)
          RNov1c_frm.sInitMovExternos
          RNov1c_frm.Combo1(1).Clear
          mCodAlfa = mFlex.Text
          'mFlex.Col = 2 'mp20160523
          'mCodAlfa = "(" & mCodAlfa & ")-" & mFlex.Text 'mp20160523
          mKm = fGetKm(mCodAlfa)
          mCodAlfa = "(" & mCodAlfa & ")-" & mKm 'mp20160523
          mFlex.Col = 3
          mCodAlfa = mCodAlfa & " " & mFlex.Text & " [" & mFlex.TextMatrix(mFlex.Row, 4) & "]"
          RNov1c_frm.Combo1(1).AddItem mCodAlfa
          RNov1c_frm.Combo1(1).ListIndex = 0
   End Select
End Sub


Private Sub mnuPMoAMoSub_Click(Index As Integer) 'MENU de Moviles para asignar
 Dim mI As Integer
 
   If Index <> 30 Then
      RNov1c_frm.Show
      RNov1c_frm.Check1.Visible = False
      If mnuPMoSub(0).Tag <> "" Then  'Es una demora
         RNov1c_frm.Frame1.Caption = "Pedido de Móvil por Demora"
         RNov1c_frm.lCodAlfa.Tag = mnuPMoSub(0).Tag  'paso fecha/hora de esa demora
         mFlex.Col = 6
         RNov1c_frm.Text1(0).Text = mFlex.Text
      Else
         RNov1c_frm.Frame1.Caption = "Pedido de Móvil"
      End If
      RNov1c_frm.lCodAlfa.Visible = True
      'mFlex.Col = 2    'mp20160523
      'mKm = mFlex.Text 'mp20160523
      'mFlex.Col = 0    'mp20160523
      mFlex.Col = 0     'mp20160523
      mKm = fGetKm(mFlex.Text)   'mp20160523
      mReferencia = fGetCodigoReferencia(mFlex.Text)  'mp20160523
      If mFlex.Text = "" Then 'Por si no tenía novedades
         mCodAlfa = fNewCodAlfa
         RNov1c_frm.lCodAlfa.Caption = mCodAlfa
         mFlex.Col = 1
         RNov1c_frm.lCodAlfa.Tag = mFlex.Text & " " & mKm
      Else
         RNov1c_frm.lCodAlfa.Caption = mFlex.Text
         sObtOrigen "", mFlex.Text, RNov1c_frm.Combo1(1)
      End If
      mFlex.Col = 3
      mSent = mFlex.Text
      fPasar_Km_Sent mKm, mSent, mFlex.TextMatrix(mFlex.Row, 4), mReferencia
      RNov1a_frm.List1.Visible = True
      RNov1a_frm.List1.AddItem mnuPMoAMoSub(Index).Tag 'En los list1.tag estan los códigos de los móviles
      RNov1a_frm.Command2(0).Visible = False
      RNov1a_frm.Label3.Visible = True
   End If
End Sub

Private Sub mnuPMoATaSub_Click(Index As Integer)
Dim mI As Integer
   If Index > 0 Then
      RNov1c_frm.Show
      RNov1c_frm.sInitTareas
      RNov1c_frm.lCodAlfa.Caption = mFlex.Text
      RNov1c_frm.lCodAlfa.Visible = True
      mFlex.Col = 2
      mKm = mFlex.Text
      mFlex.Col = 3
      mSent = mFlex.Text
      mFlex.Col = 7
      mReferencia = mFlex.Text
      fPasar_Km_Sent mKm, mSent, mFlex.TextMatrix(mFlex.Row, 4), mReferencia
      RNov1a_frm.List1.Visible = True
      RNov1a_frm.List1.AddItem mnuPMoATaSub(Index).Tag
      RNov1a_frm.Label3.Visible = True
   End If
End Sub

Private Sub mnuPMoAAcc_Click()
Dim mObj As New clRNov
  '==============================================================================
  '''''''''''''''''''''CODIGO PARA MOSTRA POP UP''''''''''''''''''''''''''''''''
  'RNov12.Show
   
  'RNov1a_frm.Enabled = False
  'RNov1b_frm.Enabled = False
  'RNov1d_frm.Enabled = False
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '==============================================================================
   
   
   'mObj.waze_enviar_accidente mFlex.TextMatrix(mFlex.Tag, 0), "UNA DESCRIPCION", 2
   'mObj.waze_asignar_movil mFlex.TextMatrix(mFlex.Tag, 0)
     
   
  '==============================================================================
  '''''''''''''''''''''CODIGO ORIGINAL DE MARCAR COMO ACCIDENTE''''''''''''''''''
   
   If MsgBox("Seguro de registrar este código como Accidente?", vbYesNo, sMessage) = vbYes Then
      If Not mObj.bExistDatoTabla("enlace", "codigo='" & mFlex.TextMatrix(mFlex.Tag, 0) & "' and fecha='" & Format(mFlex.TextMatrix(mFlex.Tag, 1), "yyyy-mm-dd") & "'") Then
         mObj.xInEnlace mFlex.TextMatrix(mFlex.Tag, 0), "", mFlex.TextMatrix(mFlex.Tag, 1), Trim(Right(MDI.mUser, 15))
         
         mObj.xUpActualizar mPc, 0
         RNov1a_frm.Label4.Visible = mObj.bExistDatoTabla("novedades2", "codnov='D'")

         
         If MsgBox("Enviar información del accidente a Waze?", vbYesNo, sMessage) = vbYes Then
            RNov11.Show
            RNov11.sInitFormCod mFlex.TextMatrix(mFlex.Tag, 0), 0

         Else 'registra en Log
         
            Unload RNov1b_frm
            Unload RNov1d_frm
            Load RNov1b_frm
            Load RNov1d_frm
            'mObj.xInLogWaze mFlex.TextMatrix(mFlex.Tag, 0), 1, "CANCELADO"
         End If

      Else
         MsgBox "Este código ya fue registrado como accidente.", vbInformation, sMessage
      End If
   End If
   Set mObj = Nothing
'==============================================================================

End Sub

Private Sub mnuPMoFEvent_Click()
   Dim mObj As New clRNov
   Dim mCod As String
   Dim mFecha As String
   Dim mWhere As String
   Dim mI As Integer
   mCod = mFlex.TextMatrix(mFlex.Tag, 0)
   
   If Not mObj.bExistDatoTabla("moviles", "ToolTip like '%" & mCod & "%'") Or mCod = "" Then

       If MsgBox("Seguro desea finalizar el evento?", vbYesNo, sMessage) = vbYes Then
          
          mWhere = "finalizado = 0 AND Codigo = '" & mCod & "'"
          If mCod = "" Then
             mFecha = mFlex.TextMatrix(mFlex.Tag, 1)
             mWhere = mWhere + " AND Fecha = '" & Format(mFecha, "yyyy-mm-dd hh:mm:ss") & "' "
          End If
          
          sMsgEspere Me, "Finalizando evento... espere.", True
            mObj.xUpNovedadesSet "finalizado = 1 ", mWhere
            Sleep (3)
          sMsgEspere Me, "", False
       
       'RNov1a_frm.recargarRNov1d
       
      
'         For mi = 1 To mFlex.Rows - 1
'
'            If mFlex.TextMatrix(mi, 0) = mCod Then
'
'               mFlex.RemoveItem (mi)
'
'            End If
'
'         Next



'         For mi = mFlex.Rows - 1 To 2 Step -1
'
'            'If mFlex.TextMatrix(mi, 0) = mCod Then
'
'               'If mi <> 1 Then
'               mFlex.RemoveItem (mi)
'               'End If
'            'End If
'
'         Next
'
'
'          mFlex.FixedRows = 1 ' o x, las que tengamos
'          mFlex.Rows = mFlex.FixedRows


         'Inicio
      
      
      'For i = 1 To .Rows - 1
      
      
       '   mFlex.RemoveItem(
       
       
          'Unload RNov1d_frm
          'Load RNov1d_frm
       
       
          'mObj.xUpActualizar mPc, 0
          'RNov1a_frm.Label4.Visible = mObj.bExistDatoTabla("novedades2", "codnov='D'")
          'Unload RNov1b_frm
          
          On Error Resume Next
          Unload RNov1d_frm
          If Err.Description <> "" Then
            MsgBox "Avisar a sistemas de este error:" & Chr(13) & Err.Description
          End If
          
          'Load RNov1b_frm
          On Error Resume Next
          Load RNov1d_frm
          If Err.Description <> "" Then
            MsgBox "Avisar a sistemas de este error:" & Chr(13) & Err.Description
          End If
          
       End If
       
    Else
       MsgBox "No es posible finalizar el evento cuando aun tiene moviles o gruas asignadas", vbInformation, sMessage
    End If
   
   Set mObj = Nothing
End Sub
Private Sub mnuPMoAMB_Click()
   RNov1c_frm.Show
   RNov1c_frm.Frame1.Caption = "Pedido de Móviles"
   RNov1c_frm.Label2.Caption = Right(mnuPMoAMB.Caption, 4)
   RNov1c_frm.sInitMovExternos
   RNov1c_frm.Combo1(1).Clear
   mCodAlfa = mFlex.Text
   mFlex.Col = 2
   mCodAlfa = "(" & mCodAlfa & ")-" & mFlex.Text
   mFlex.Col = 3
   mCodAlfa = mCodAlfa & " " & mFlex.Text & " " & mFlex.TextMatrix(mFlex.Row, 4)
   RNov1c_frm.Combo1(1).AddItem mCodAlfa
   RNov1c_frm.Combo1(1).ListIndex = 0
End Sub

'*****************************************************************
'------  FIN MENU ------------------------------------------------
'*****************************************************************

Private Sub Inicio()
Dim mObj As New clRNov
Dim mI As Integer
Dim mCont As Integer
Dim Cod, CodAnterior, xSent, mMov As String
Dim xCodRamal As String
Dim xEsAccidente As Boolean
  
   'Me.Width = 15300
   'Me.Height = 6600
   Me.Width = 19695
   Me.Height = 7350
   Me.Top = RNov1b_frm.Height + RNov1a_frm.Height + 10
   Me.Left = 0
   With mFlex
      .ColWidth(0) = 1100
      .ColWidth(1) = 2000
      .ColWidth(2) = 600
      .ColWidth(3) = 400
      .ColWidth(4) = 450
      .ColWidth(5) = 450
      .ColWidth(6) = 14700
      .ColWidth(7) = 0
      
      .Row = 0
      .Font = "Arial"
         .TextMatrix(0, 0) = "Código"
         .TextMatrix(0, 1) = "Fecha/Hora"
         .TextMatrix(0, 2) = "Km"
         .TextMatrix(0, 3) = "Sen"
         .TextMatrix(0, 4) = "Ram"
         .TextMatrix(0, 5) = "Mov"
         .TextMatrix(0, 6) = "Novedad"
         .TextMatrix(0, 7) = ""
         .FixedAlignment(6) = 0
      For mI = 0 To 7
         .Col = mI
         .CellFontBold = True
         .FixedAlignment(mI) = 4
      Next
      .MergeCol(0) = True
   End With
   mI = 0


   If RNov1a_frm.verTodos Then
      mCont = mObj.iCountNov2()
      Set mRec = mObj.oTablaDina("novedades2", "order by fecha desc limit 200")
   Else
      mCont = mObj.iCountNovPendientes()
      Set mRec = mObj.oTablaDina("novedades2 where finalizado = 0", "order by fecha desc limit 200")
   End If
   
   If mCont > 200 Then
      mCont = 200
   End If
      
   If Not mRec.EOF Then
      mRec.MoveLast
   End If


   'Sleep (1)
   
   CodAnterior = "valorForzado"
   Do While Not mRec.BOF
      mFlex.Font = "Arial"
      mMov = ""
      If VarType(mRec!Codigo) = 1 Then
         Cod = ""
      Else
         Cod = mRec!Codigo
      End If
      
      'xEsAccidente = True
      
      If VarType(mRec!sent) = 1 Or mRec!sent = 0 Then
         xSent = ""
         xCodRamal = ""
      Else
         xSent = Left(mObj.sTablaDescr("sentidos", "codigo=" & mRec!sent, 1), 2)
         xCodRamal = Mid(mObj.sTablaDescr("ramales", "codigo=" & mRec!codramal, 2), 2, 2)
      End If
      If mRec!CodNov = "A" Or mRec!CodNov = "C" Or mRec!CodNov = "E" Or mRec!CodNov = "E" Or mRec!CodNov = "MM" Then
         mMov = "SI"
      End If
      
      
      
      If Cod <> "" Then
         If Cod <> CodAnterior Then
            If mObj.bExistDatoTabla("enlace", "codigo='" & Cod & "' and fecha='" & Format(mRec!Fecha, "yyyy-mm-dd") & "'") Then
               xEsAccidente = True
            Else
               xEsAccidente = False
            End If
         End If
      Else
         xEsAccidente = False
      End If
      CodAnterior = Cod
      
      If mI = (mCont - 3) Then
         sCargar Cod, mRec!Fecha, mRec!km, xSent, mMov, mRec!descripcion, mFlex, True, xCodRamal, mRec!codreferencia, xEsAccidente
      Else
         sCargar Cod, mRec!Fecha, mRec!km, xSent, mMov, mRec!descripcion, mFlex, False, xCodRamal, mRec!codreferencia, xEsAccidente
         mI = mI + 1
      End If
      If mRec!CodNov = "D" Then
         mFlex.Col = 5
         Set mFlex.CellPicture = LoadPicture(App.Path & "\Regnovedades\Image\Reloja.gif")
         RNov1a_frm.Label4.Visible = True
      End If
      mRec.MovePrevious
   Loop
   mRec.Close
   RNov1d_frm.Refresh
   Set mObj = Nothing
End Sub

Public Function sCargar(ByVal xCod As String, xFecha As Date, ByVal xKm As Double, ByVal xSent As String, ByVal xMov As String, ByVal xNov As String, xFlex As Object, ByVal mBold As Boolean, ByVal xCodRamal As String, ByVal xCodReferencia As String, xEsAccidente As Boolean)
Dim mColor As Boolean
Dim mFlag As Boolean
Dim mI As Integer
Dim mRow As Integer
Dim mColour As Double
Dim Kolor As Double
 
   mColor = False
   mFlag = False
   If xFlex.Rows = 1 Then 'Es el primero
      mColour = &HC0FFFF 'amarillo
      xFlex.AddItem xCod & vbTab & xFecha & vbTab & Format(xKm, "#0.00") & vbTab & xSent & vbTab & xCodRamal & vbTab & xMov & vbTab & xNov & vbTab & xCodReferencia
      xFlex.Row = 1
      For mI = 0 To xFlex.Cols - 1
         xFlex.Col = mI
         xFlex.CellBackColor = mColour
         
         If xEsAccidente Then
               xFlex.CellForeColor = &HFFF 'rojo
         Else
               xFlex.CellForeColor = &H0
         End If
         
         
         
      Next
   Else
      xFlex.Col = 0
      For mI = 1 To xFlex.Rows - 1
         xFlex.Row = mI
         If xFlex.Text = xCod Then
            mRow = mI
            mFlag = True
         End If
      Next
      If mFlag And xCod <> "" Then
         xFlex.Row = mRow
         xFlex.AddItem xCod & vbTab & xFecha & vbTab & Format(xKm, "#0.00") & vbTab & xSent & vbTab & xCodRamal & vbTab & xMov & vbTab & xNov & vbTab & xCodReferencia, mRow + 1
         mColour = xFlex.CellBackColor
         xFlex.MergeRow(mRow + 1) = True
         xFlex.Col = 0
         xFlex.Row = mRow + 1
         For mI = 0 To xFlex.Cols - 1
            xFlex.Col = mI
            xFlex.CellBackColor = mColour
            
            
            If xEsAccidente Then
               xFlex.CellForeColor = &HFFF 'rojo
            Else
               xFlex.CellForeColor = &H0
            End If

         Next
      Else
         xFlex.Row = 1
         xFlex.Col = 1
         mColour = &HC0FFFF
         Kolor = xFlex.CellBackColor
         If Kolor = mColour Then
            mColour = &HFFFFFF 'Blanco
         End If
         xFlex.Col = 0
         xFlex.AddItem xCod & vbTab & xFecha & vbTab & Format(xKm, "#0.00") & vbTab & xSent & vbTab & xCodRamal & vbTab & xMov & vbTab & xNov & vbTab & xCodReferencia, 1
         xFlex.Row = 1
         For mI = 0 To xFlex.Cols - 1
            xFlex.Col = mI
            xFlex.CellBackColor = mColour
           
            
            If xEsAccidente Then
               xFlex.CellForeColor = &HFFF 'rojo
            Else
               xFlex.CellForeColor = &H0
            End If
            
            
            'xFlex.ForeColor = &HFF0101
           'If xFlex.Row Mod 2 = 0 Then
            'xFlex.ForeColor = &HFF0101
            'End If
            
            'xFlex.ForeColor = &HFFF 'rojo
           
            
         Next
         
          'xFlex.ForeColor = &H0
         
      End If
   End If
   xFlex.Col = 0
   xFlex.CellAlignment = 4
   xFlex.Col = 3
   xFlex.CellAlignment = 4
   xFlex.Col = 4
   xFlex.CellAlignment = 4
   
   xFlex.Col = 7
   xFlex.CellAlignment = 4
   If mBold Then
      For mI = 1 To xFlex.Cols - 1
         xFlex.Col = mI
         xFlex.CellFontBold = 1
         xFlex.CellForeColor = &HFF0000
         
         If xEsAccidente Then
               xFlex.CellForeColor = &HFFF 'rojo
         Else
               xFlex.CellForeColor = &HFF0000
         End If
         
      Next
   End If
'   If xEsAccidente Then
'      xFlex.ForeColor = &HFFF 'rojo
'   Else
'      xFlex.ForeColor = &H0
'   End If
   
End Function

Private Function fMovDisponibles(pObj As Object, pImg1 As String, pImg2 As String, pImg3 As String) As Boolean
Dim mI As Integer
Dim mJ As Integer
Dim mFlag As Boolean
 
   mFlag = False
   For mI = 0 To pObj.UBound
      If pObj(mI).Tag = pImg1 Or pObj(mI).Tag = pImg2 Then
          For mJ = 0 To mnuPMoAMoSub.UBound - 1
             If pObj(mI).Caption = mnuPMoAMoSub(mJ).Caption Then
                mnuPMoAMoSub(mJ).Visible = True
                mJ = 999
                mFlag = True
             End If
          Next
      Else
          If pObj(mI).Tag <> pImg3 Then
             For mJ = 1 To mnuPMoATaSub.UBound
                If pObj(mI).Caption = mnuPMoATaSub(mJ).Caption Then
                   mnuPMoATaSub(mJ).Visible = True
                   mnuPMoATaSub(0).Visible = False
                   mJ = 999
                End If
             Next
          End If
      End If
   Next
   fMovDisponibles = mFlag
End Function

Private Sub sInitMenuDemora(pFlag As Boolean)
Dim mI As Integer
   mnuPMoSub(1).Visible = pFlag
   mnuPMoSub(2).Visible = pFlag
   For mI = 5 To 8
      mnuPMoSub(mI).Visible = pFlag
   Next
   mnuPMoATar.Visible = pFlag
End Sub

