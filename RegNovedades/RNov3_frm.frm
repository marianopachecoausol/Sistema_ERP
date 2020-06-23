VERSION 5.00
Begin VB.Form RNov3_frm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7035
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
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
   ScaleHeight     =   3465
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3295
      Left            =   80
      TabIndex        =   6
      Top             =   120
      Width           =   6900
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   992
         Width           =   4700
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   400
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Doble código de parte"
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
         Height          =   240
         Left            =   2280
         TabIndex        =   10
         Top             =   1540
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         MaxLength       =   150
         TabIndex        =   1
         Top             =   1585
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   5640
         TabIndex        =   5
         Top             =   2680
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Grabar"
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
         Left            =   4200
         MaskColor       =   &H8000000B&
         TabIndex        =   4
         Top             =   2560
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2020
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   0
         Top             =   1585
         Width           =   735
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Top             =   2020
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ref."
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
         Index           =   3
         Left            =   1440
         TabIndex        =   13
         Top             =   1067
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ramal"
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
         Index           =   2
         Left            =   1440
         TabIndex        =   12
         Top             =   520
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Left            =   240
         TabIndex        =   9
         Top             =   2620
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sentido"
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
         Left            =   1320
         TabIndex        =   8
         Top             =   2065
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Km"
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
         Left            =   1680
         TabIndex        =   7
         Top             =   1615
         Width           =   270
      End
   End
End
Attribute VB_Name = "RNov3_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mRec As New ADODB.Recordset
Public xObjGrua As Object
Dim mResp As Boolean
Dim mPc As String

Private Sub Form_Load()
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
   mPc = Mid(MDI.mPCname, 1, Len(MDI.mPCname) - 1)
   Me.Width = 7125
   Me.Height = 3550
   sAlinearForm Me
   RNov1a_frm.Enabled = False
   RNov1b_frm.Enabled = False
   RNov1d_frm.Enabled = False
   Set mRec = mObj.oTabla("ramales", "")
   Do While Not mRec.EOF
     Combo1(1).AddItem mRec!Descripcion & Space(60) & mRec!Codigo
     mRec.MoveNext
   Loop
   mRec.Close
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Private Sub Combo1_Click(Index As Integer)
   Select Case Index
      Case 1
         If Combo1(1).ListIndex >= 0 Then
            sLlenoSentido
            sLlenoReferencia
         End If
   
      Case 2
         If Combo1(2).ListIndex >= 0 Then
            Dim mCodReferencia As String
            mCodReferencia = Right(Combo1(2).Text, 3)
            sCompletaKM mCodReferencia
         End If
   End Select
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mObj As New clRNov
Dim mCodAlfa As String
Dim mKm As String
Dim mSent As String
Dim mCodRamal As String
Dim mCodReferencia As String
Dim mDescr As String
Dim mImgNov As String
Dim mNroParte As Integer
Dim mCodNextel As String
Dim mError As Boolean
Dim mClima As String

   If Combo1(2).ListIndex = -1 Then
      mCodReferencia = 0 'Referencia nula
   Else
      mCodReferencia = Trim(Right(Combo1(2).Text, 3))
   End If

   If Index = 0 Then
      mError = False
      Select Case Frame1.Tag
         Case "QTH"
            'mKm = Format(Text1.Text, "00.00")
            mKm = Text1.Text 'mp 20160309
            mCodAlfa = ""
            If Combo1(0).ListIndex > -1 Then
               If Progresiva_Ok(mKm, Trim(Right(Combo1(0).Text, 2))) = True Then
                  mDescr = Right(Frame1.Caption, 4) & " - QTH"
                  If Mid(Frame1.Caption, 24, 1) = "G" Then
                     RNov1b_frm.Grua(RNov1b_frm.mIndexMov).Picture = LoadPicture(App.Path & "\RegNovedades\image\iconos\g4.gif")
                     RNov1b_frm.Label4(RNov1b_frm.mIndexMov).Tag = "g4"
                     mObj.xUpToolMov "Codigo='" & Right(Frame1.Caption, 4) & "'", "CodNov='g4'"
                     mResp = mObj.xInsNovedades("", mCodAlfa, Trim(Right(MDI.mUser, 15)), mKm, Right(Combo1(0).Text, 2), "SSV", mDescr, "J", 0, Right(Frame1.Caption, 4), "", 0, "", "", 0, "", "", 0, "", "", Right(Combo1(1).Text, 1), mCodReferencia)
                     mObj.xUpActualizarNot mPc, 1
                  Else
                     If Mid(Frame1.Caption, 12, 3) = "AMB" Then
                        RNov1b_frm.MovExternos(RNov1b_frm.mIndexMov).Picture = LoadPicture(App.Path & "\RegNovedades\image\iconos\a0.gif")
                        RNov1b_frm.Label8(RNov1b_frm.mIndexMov).Tag = "a0"
                        mResp = mObj.xInsNovedades("", mCodAlfa, Trim(Right(MDI.mUser, 15)), mKm, Right(Combo1(0).Text, 2), "SSV", mDescr, "QT", 0, RNov1b_frm.MovExternos(RNov1b_frm.mIndexMov).Tag, "", 0, "", "", 0, "", "", 0, "", "", Right(Combo1(1).Text, 1), mCodReferencia)
                        mObj.xUpActualizarNot mPc, 1
                        mObj.xUpToolMov "Codigo='" & RNov1b_frm.MovExternos(RNov1b_frm.mIndexMov).Tag & "'", "CodNov='a0'"
                     Else
                        RNov1b_frm.Pat(RNov1b_frm.mIndexMov).Picture = LoadPicture(App.Path & "\RegNovedades\image\iconos\p1.gif")
                        RNov1b_frm.Label3(RNov1b_frm.mIndexMov).Tag = "p1"
                        mResp = mObj.xInsNovedades("", mCodAlfa, Trim(Right(MDI.mUser, 15)), mKm, Right(Combo1(0).Text, 2), "SSV", mDescr, "QT", 0, Right(Frame1.Caption, 4), "", 0, "", "", 0, "", "", 0, "", "", Right(Combo1(1).Text, 1), mCodReferencia)
                        mObj.xUpActualizarNot mPc, 1
                        mObj.xUpToolMov "Codigo='" & Right(Frame1.Caption, 4) & "'", "CodNov='p1'"
                     End If
                  End If
                  Set mObj = Nothing
                  Unload RNov3_frm
                  Unload RNov1d_frm
                  RNov1d_frm.Show
               Else
                  mError = True
                  Set mObj = Nothing
               End If
            Else
               MsgBox "Falta comletar datos", vbCritical, sMessage
            End If
            
         Case "LIB"
            If Combo1(0).ListIndex > -1 Then
               'mKm = Mid(Left(xObjGrua(RNov1b_frm.mIndexMov).ToolTipText, InStr(1, xObjGrua(RNov1b_frm.mIndexMov).ToolTipText, " ") - 1), 11) 'Km 'mp 20160523
               mSent = Left(Right(xObjGrua(RNov1b_frm.mIndexMov).ToolTipText, 7), 2) 'mp 20160315
               mCodRamal = Right(xObjGrua(RNov1b_frm.mIndexMov).ToolTipText, 4) 'mp 20160315
               mCodRamal = mObj.sTablaDescr("ramales", "abrevia='" & mCodRamal & "'", 0) 'obtengo código de tabla
               mSent = mObj.sTablaDescr("sentidos", "left(descripcion,2)='" & mSent & "' and codramal=" & mCodRamal, 0) 'obtengo código de tabla
               mCodReferencia = fGetCodigoReferencia(Trim(Mid(xObjGrua(RNov1b_frm.mIndexMov).ToolTipText, 2, 7)))
               mKm = fGetKm(Trim(Mid(xObjGrua(RNov1b_frm.mIndexMov).ToolTipText, 2, 7)))
               Set mRec = mObj.oTabla("ultimos", "")
               mNroParte = 0
               If Not mRec.EOF Then
                  mNroParte = mRec!nroparte
                  If Check1.Value = 1 Then '2 cod. partes de servicios
                     mObj.xUpUltimos (mNroParte + 1)
                     MsgBox "Número de Parte: " & mNroParte & " y " & (mNroParte + 1), vbInformation, "Servicio de Grúa."
                     mDescr = "Móvil " & Right(Frame1.Caption, 4) & " Liberado.  -   " & Trim(Left(Combo1(0).Text, 3)) & "   -   Nro. Parte: " & mNroParte & " y " & (mNroParte + 1)
                  Else
                     mObj.xUpUltimos mNroParte
                     MsgBox "Número de Parte: " & mNroParte, vbInformation, "Servicio de Grúa."
                     mDescr = "Móvil " & Right(Frame1.Caption, 4) & " Liberado.  -   " & Trim(Left(Combo1(0).Text, 3)) & "   -   Nro. Parte: " & mNroParte
                  End If
               End If
               mRec.Close
               If Left(xObjGrua(RNov1b_frm.mIndexMov).Tag, 2) = "GP" Then
                  mImgNov = "x1"
                  RNov1b_frm.Label7(RNov1b_frm.mIndexMov).Tag = mImgNov
               Else
                  mImgNov = "g1"
                  RNov1b_frm.Label4(RNov1b_frm.mIndexMov).Tag = mImgNov
               End If
              
              mObj.xUpEstMoviles xObjGrua(RNov1b_frm.mIndexMov).Tag, "L", "", mImgNov
              xObjGrua(RNov1b_frm.mIndexMov).ToolTipText = ""
              xObjGrua(RNov1b_frm.mIndexMov).Picture = LoadPicture(App.Path & "\RegNovedades\image\iconos\" & mImgNov & ".gif")
              If Check1.Value = 1 Then '2 cod. partes de servicios
                 mResp = mObj.xInsNovedades("", Label2.Caption, Trim(Right(MDI.mUser, 15)), mKm, mSent, "SSV", mDescr, "TT", "0", Right(Frame1.Caption, 4), Trim(Left(Combo1(0).Text, 3)), mNroParte, "", "", (mNroParte + 1), "", "", 0, "", "", mCodRamal, mCodReferencia)
              Else
                 mResp = mObj.xInsNovedades("", Label2.Caption, Trim(Right(MDI.mUser, 15)), mKm, mSent, "SSV", mDescr, "TT", "0", Right(Frame1.Caption, 4), Trim(Left(Combo1(0).Text, 3)), mNroParte, "", "", 0, "", "", 0, "", "", mCodRamal, mCodReferencia)
              End If
              mObj.xUpActualizarNot mPc, 1
              Set xObjGrua = Nothing
              Unload RNov1d_frm
              Load RNov1d_frm
              Unload RNov2_frm
           Else
              mError = True
           End If
           Set mObj = Nothing
           
         Case "OPER"
            If Progresiva_Ok(Trim(Text1.Text), Trim(Right(Combo1(0).Text, 2))) And Combo1(0).ListIndex > -1 Then
               'mKm = Format(Trim(Text1.Text), "00.00")
               mKm = Trim(Text1.Text) 'mp 20160309
               mClima = ClimaOK(mKm)
               mResp = mObj.xInsNovedades("", "", Trim(Right(MDI.mUser, 15)), mKm, Right(Combo1(0).Text, 2), "SSV", "Operativo " & Right(Frame1.Caption, 4), "P", mClima, Right(Frame1.Caption, 4), "", 0, "", "", 0, "", "", 0, "", "", Right(Combo1(1).Text, 2), mCodReferencia)
               mObj.xUpActualizarNot mPc, 1
               Unload RNov1d_frm
               RNov1d_frm.Show
            Else
               mError = True
            End If
            Set mObj = Nothing
           
         Case "RADIO"
            If Trim(Text2.Text) <> "" And Combo1(0).ListIndex > -1 Then
               mClima = ClimaOK(19.5)
               mResp = mObj.xInsNovedades("", "", Trim(Right(MDI.mUser, 15)), "19.5", "1", "SSV", Combo1(0).Text & Space(15) & " - " & Trim(Text2.Text), "F", mClima, "", "", 0, "", "", 0, "", "", 0, "", "", "1", 26)
               mObj.xUpActualizarNot mPc, 1
            Else
               mError = True
            End If
            Set mObj = Nothing
            
         Case "COMB"
            If Trim(Text1.Text) <> "" And Trim(Text3.Text) <> "" Then
               mObj.xInsCombustible Now(), Right(Frame1.Caption, 4), Trim(Text3.Text), Trim(Text1.Text)
            Else
               mError = True
            End If
            Set mObj = Nothing
      End Select
   Else
      Unload RNov3_frm
      Set mObj = Nothing
   End If
   If Not mError Then
      RNov1a_frm.Enabled = True
      RNov1b_frm.Enabled = True
      Unload RNov1d_frm
      RNov1d_frm.Show
      Unload RNov3_frm
   Else
      MsgBox "Existe un Error", vbCritical, sMessage
   End If
End Sub

Public Sub sInitQTH()
Frame1.Tag = "QTH"
Label1(2).Visible = True
Label1(3).Visible = True
Combo1(2).Visible = True
Combo1(1).Visible = True
Combo1(0).Visible = True
Combo1(0).Width = 4000

End Sub

Public Sub sInitFreeGrua()
Dim mObj As New clRNov
Label1(0).Visible = False
Text1.Visible = False
Label1(1).Left = Label1(1).Left - 960
Label1(1).Caption = "Servicios"
Label2.Visible = True
Combo1(0).Left = Combo1(0).Left - 960
Combo1(0).Width = 4000
Combo1(0).Visible = True
Combo1(1).Visible = False
Combo1(2).Visible = False
Set mRec = mObj.oTabla("servicios", "WHERE fecha_baja IS NULL or fecha_baja = '0000-00-00 00:00:00'")
Do While Not mRec.EOF
   Combo1(0).AddItem mRec.Fields(0) & " - " & mRec.Fields(1)
   mRec.MoveNext
Loop
mRec.Close
Check1.Visible = True
Me.SetFocus
Set mObj = Nothing
End Sub

Public Sub sInitRadio()
Dim mObj As New clRNov
   Frame1.Caption = "Novedad de Emisora"
   Label1(0).Left = 600
   Label1(0).Caption = "Novedad"
   Label1(1).Left = 150
   Label1(1).Caption = "Emisora"
   Text1.Visible = False
   Text2.Visible = True
   Text2.Left = 920
   Text2.Width = 5500
   Combo1(0).Left = 920
   Combo1(0).Width = 4300
   Combo1(0).Visible = True
   Combo1(1).Visible = False
   Combo1(2).Visible = False
   Set mRec = mObj.oTabla("emisoras", "where fecha_baja IS NULL or fecha_baja = '0000-00-00 00:00:00'")
   Do While Not mRec.EOF
      Combo1(0).AddItem mRec.Fields(0) & " - " & mRec.Fields(1)
      mRec.MoveNext
   Loop
   mRec.Close
   Text2.SetFocus
   Set mObj = Nothing
End Sub

Public Sub sInitCombustible()
   Frame1.Tag = "COMB"
   Label1(0).Left = 1600
   Label1(0).Caption = "Litros"
   Label1(1).Left = 1300
   Label1(1).Caption = "Kilómetros"
   Text1.Visible = True
   Text3.Visible = True
   Combo1(0).Visible = False
   Combo1(1).Visible = False
   Combo1(2).Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   RNov1a_frm.Enabled = True
   RNov1b_frm.Enabled = True
   RNov1d_frm.Enabled = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 46 Then
      KeyAscii = fNumeroKeyPress(KeyAscii)
   End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
   KeyAscii = fUcaseKeyPress(KeyAscii)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 46 Then
      KeyAscii = fNumeroKeyPress(KeyAscii)
   End If
End Sub

Private Sub sLlenoSentido()
Dim mCodRamal As String
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
   
   mCodRamal = Right(Combo1(1).Text, 1)
   Combo1(0).Clear
   Set mRec = mObj.oTabla("sentidos", "where codRamal = " & mCodRamal & " order by 2")
   Do While Not mRec.EOF
     Combo1(0).AddItem mRec!Descripcion & Space(60) & mRec!Codigo
     mRec.MoveNext
   Loop
   mRec.Close
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Private Sub sLlenoReferencia()
Dim mCodRamal As String
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
   
   mCodRamal = Right(Combo1(1).Text, 1)
   Combo1(2).Clear
   Set mRec = mObj.oTabla("referencias", "where codRamal = " & mCodRamal & " order by 2")
   Do While Not mRec.EOF
     Combo1(2).AddItem mRec!Descripcion & Space(100) & mRec!Codigo
     mRec.MoveNext
   Loop
   mRec.Close
   Set mObj = Nothing
   Set mRec = Nothing
End Sub
Private Sub sCompletaKM(pCodReferencia As String)
   Dim mObj As New clRNov
   Text1.Text = mObj.sTablaDescr("referencias", "codigo=" & pCodReferencia, 2)
   Set mObj = Nothing
End Sub

