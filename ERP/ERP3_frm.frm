VERSION 5.00
Begin VB.Form ERP3_frm 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3465
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5160
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "ERP3_frm.frx":0000
   ScaleHeight     =   3465
   ScaleWidth      =   5160
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E8E8E3&
      Caption         =   "Cancelar"
      Height          =   315
      Index           =   1
      Left            =   2820
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E8E8E3&
      Caption         =   "Aceptar"
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
      Height          =   315
      Index           =   0
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H8000000D&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2100
      MaxLength       =   25
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1800
      Width           =   1755
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H8000000D&
      Height          =   285
      Index           =   0
      Left            =   2100
      MaxLength       =   25
      TabIndex        =   2
      Top             =   1440
      Width           =   1755
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Requisitos"
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   3960
      MouseIcon       =   "ERP3_frm.frx":126F
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E8E8E3&
      Height          =   195
      Index           =   3
      Left            =   60
      TabIndex        =   7
      Top             =   3180
      Width           =   75
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AUTOPISTAS DEL SOL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E8E8E3&
      Height          =   195
      Index           =   2
      Left            =   2925
      TabIndex        =   6
      Top             =   3180
      Width           =   2025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   195
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   1860
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   195
      Index           =   0
      Left            =   1260
      TabIndex        =   0
      Top             =   1500
      Width           =   660
   End
End
Attribute VB_Name = "ERP3_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mObj As New clLogUser
Dim mContIntentos As Integer
Dim mCambioClave As Boolean
Dim mRec As New ADODB.Recordset

Private Sub Form_Load()
   sAlinearForm Me
   mContIntentos = 0
   Label1(3).Caption = Mid(sMessage, 25, Len(sMessage) - 24)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mInd As Integer
   If Index = 0 Then
      If Text1(0).Text <> "" And Text1(1).Text <> "" Then
         If fValidUser Then
            If Not mCambioClave Then
               Set mRec = mObj.oSistemasUsuario(Trim(Text1(0).Text))
               Do While Not mRec.EOF
                  On Error Resume Next
                  If IsNumeric(Left(ERP1_frm.Image1(mRec!codsistema).Tag, 1)) Then
                     ERP1_frm.Command1(Left(ERP1_frm.Image1(mRec!codsistema).Tag, 2)).MousePointer = 0
                  End If
                  ERP1_frm.Image1(mRec!codsistema).MousePointer = 99
                  ERP1_frm.Image1(mRec!codsistema).MouseIcon = LoadPicture(App.Path & "\ERP\Imagenes\Hand.cur")
                  mRec.MoveNext
               Loop
               mRec.Close
            End If
            ERP1_frm.Show
            MDI.BackColor = &H808080
            Unload ERP3_frm
         End If
      Else
         MsgBox "Faltan Completar Datos", vbCritical, sMessage
      End If
   Else
      Unload ERP3_frm
      Unload MDI
   End If
End Sub

Private Sub Label2_Click()
MsgBox "REQUISITOS para una clave:" & Chr(13) & Chr(13) & "Debe estar compuesta por caracteres Alfanuméricos." & Chr(13) & "Debe contener como  mínimo 2 (dos) números." _
       & Chr(13) & "Cambiarse obligatoriamente en un período máximo de 60 días." & Chr(13) & "Ser distintas por lo menos de las últimas 6 anteriores." _
       & Chr(13) & "Deben tener un tiempo mínimo de vida de 15 días.", vbInformation, sMessage
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 0 Then
      If KeyAscii <> 8 And KeyAscii <> 32 And Not (KeyAscii >= 65 And KeyAscii <= 90) And Not (KeyAscii >= 97 And KeyAscii <= 122) Then
         KeyAscii = 0
      Else
          If KeyAscii >= 65 And KeyAscii <= 90 Then
              KeyAscii = KeyAscii + 32
          End If
      End If
   Else
      KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
   End If
End Sub

Private Function fValidUser() As Boolean 'Valida usuario existente y que no esté bloqueado por 15 minutos
   fValidUser = False
   mCambioClave = False
   If Not mObj.bUserBloqueado(Trim(Text1(0).Text)) Then
      If mContIntentos < 6 Then 'INTENTOS DE USUARIO Y CLAVE
         Set mRec = mObj.oValidUsuario(Trim(Text1(0).Text), Trim(Text1(1).Text))
         mContIntentos = mContIntentos + 1
         If Not mRec.EOF Then
            MDI.mUser = mRec!apellido & ", " & mRec!nombres & Space(40) & Trim(Text1(0).Text)
            MDI.mClave = Trim(Text1(1).Text)
            If mRec!chgpass = "0" Then 'EL SISTEMA REQUIERE UN CAMBIO DE CLAVE
               mCambioClave = True
               MsgBox "Usuario nuevo en el sistema." & Chr(13) & Chr(13) & "Para entrar al sistema primero deberá cambiar la clave actual." & Chr(13) & "(USUARIOS - Cambio de Clave)", vbInformation, sMessage
            Else
               If mObj.bControlClave(Trim(Text1(0).Text), Trim(Text1(1).Text)) Then
                  MsgBox "La clave ha vencido, tiene más de 60 días de vigencia." & Chr(13) & Chr(13) & "Para entrar al sistema primero deberá cambiar la clave actual." & Chr(13) & "(USUARIOS - Cambio de Clave)", vbInformation, sMessage
                  mCambioClave = True
               End If
            End If
            fValidUser = True
         Else
            If mContIntentos = 5 Then
               If mObj.xSetDelayUser(Trim(Text1(0).Text)) Then 'ver de manejarlo con la tabla
                  MsgBox "El Usuario inhabilitado por 15 minutos por intentos fallidos reiterados", vbCritical, sMessage
               End If
            Else
               MsgBox "Verifique los Datos Ingresados", vbCritical, sMessage
            End If
         End If
         mRec.Close
      End If
   Else
      MsgBox "Usuario inhabilitado por 15 minutos por intentos fallidos reiterados", vbCritical, sMessage
   End If
End Function
