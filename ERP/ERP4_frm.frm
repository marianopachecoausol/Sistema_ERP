VERSION 5.00
Begin VB.Form ERP4_frm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Clave - "
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6510
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   4305
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000A&
         Height          =   975
         Left            =   1200
         TabIndex        =   9
         Top             =   3120
         Width           =   4095
         Begin VB.CommandButton Command1 
            Caption         =   "&Cancelar"
            Height          =   495
            Index           =   1
            Left            =   2280
            TabIndex        =   11
            Top             =   280
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Aceptar"
            Default         =   -1  'True
            Height          =   495
            Index           =   0
            Left            =   600
            TabIndex        =   10
            Top             =   280
            Width           =   1095
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2500
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1540
         Width           =   1815
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000A&
         Height          =   735
         Left            =   290
         TabIndex        =   1
         Top             =   120
         Width           =   5775
         Begin VB.Image Image1 
            Height          =   495
            Left            =   120
            Picture         =   "ERP4_frm.frx":0000
            Stretch         =   -1  'True
            Top             =   150
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cambio de Clave"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   300
            Left            =   1800
            TabIndex        =   2
            Top             =   240
            Width           =   2025
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requisitos para una Clave"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   4380
         MouseIcon       =   "ERP4_frm.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   2880
         Width           =   1860
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   960
         Width           =   4815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirmación:"
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
         Left            =   1320
         TabIndex        =   7
         Top             =   2545
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clave Nueva:"
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
         Left            =   1320
         TabIndex        =   4
         Top             =   2080
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Clave Actual:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   1320
         TabIndex        =   3
         Top             =   1580
         Width           =   1155
      End
   End
End
Attribute VB_Name = "ERP4_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Me.Height = 4740
   Me.Width = 6625
   sAlinearForm Me
   Me.Caption = Me.Caption & " " & sMessage
   ERP1_frm.Visible = False
   Label3.Caption = Trim(Left(MDI.mUser, 35))
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ERP1_frm.Visible = True
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mObj As New clLogUser
Dim mRec As New ADODB.Recordset
Dim mClave As String
Dim mInicia As Boolean
Dim mFlag As Boolean
Dim mI As Integer
Dim mJ As Integer
mInicia = False
   If Index = 0 Then
      If Text1(0).Text <> "" And Text1(1).Text <> "" And Text1(2).Text <> "" Then
         If Len(Trim(Text1(1).Text)) > 5 Then  'Control de clave mayor a 6 caracteres
            If fControlAlfa Then
               Set mRec = mObj.oValidUsuario(Trim(Right(MDI.mUser, 20)), Trim(Text1(0).Text))
               If Not mRec.EOF Then
                  If Trim(Text1(1).Text) = Trim(Text1(2).Text) Then
                     mClave = Trim(Text1(1).Text)
                     If mObj.bRepetClave(Trim(Right(MDI.mUser, 20)), mClave) Then  ' ### Control de Clave repetida
                        MsgBox "La clave ingresada debe ser diferente" & Chr(13) & " a las últimas seis utilizadas anteriormente.", vbCritical, sMessage
                     Else
                        If mObj.bDurMinClave(Trim(Right(MDI.mUser, 20))) Then   ' ### Control de Clave  duración Minima de 15 días
                           MsgBox "Imposible realizar el cambio de clave. " & Chr(13) & "La duración mínima de clave actual debe ser de 15 días", vbInformation, sMessage
                        Else
                           If mObj.bControlClave(Trim(Right(MDI.mUser, 20)), MDI.mClave) Then     'control para ver si es una cambio de clave por vencimiento.
                              mFlag = mObj.xUpClavesUsuarios(Trim(Right(MDI.mUser, 20)), NVL(mRec!CLAVE, ""), NVL(mRec!Clave1, ""), NVL(mRec!Clave2, ""), NVL(mRec!Clave3, ""), NVL(mRec!Clave4, ""))
                           End If
                           mFlag = mObj.xUpChgClaveUser(Trim(Right(MDI.mUser, 20)), mClave)
                           'mObj.xInsertMD5 Trim(Right(MDI.mUser, 20)), mClave
                           MsgBox "Clave Cambiada con Exito!", vbInformation, sMessage
                           Text1(0).Text = ""
                           Text1(1).Text = ""
                           Text1(2).Text = ""
                           If mRec!chgpass = "0" Then
                              MsgBox "El Sistema se Cerrará, luego deberá Entrar con la Nueva Clave", vbInformation, sMessage
                              mInicia = True
                           End If
                        End If
                     End If
                  Else
                     MsgBox "Clave de Confirmación Erronea", vbCritical
                  End If
               Else
                  MsgBox "Clave de Actual Erronea", vbCritical
               End If
               mRec.Close
            End If
         Else
            MsgBox "El Mínimo de Caracteres de la Clave debe ser 6 (Seis)", vbCritical, "Sistema Global"
         End If
      Else
         MsgBox "Faltan Ingresar Datos!", vbCritical, "Sistema Global"
      End If
      If mInicia Then
         Unload ERP4_frm
         Unload MDI
         MDI.Show
      End If
   Else
      mExitSist 9
      ShowMenu 9, False, True
      Unload ERP4_frm
   End If
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Private Sub Label4_Click()
   MsgBox "REQUISITOS para una clave:" & Chr(13) & Chr(13) & "Debe estar compuesta por caracteres Alfanuméricos." & Chr(13) & "Debe contener como  mínimo 2 (dos) números." _
       & Chr(13) & "Cambiarse obligatoriamente en un período máximo de 60 días." & Chr(13) & "Ser distintas por lo menos de las últimas 6 anteriores." _
       & Chr(13) & "Deben tener un tiempo mínimo de vida de 15 días." & Chr(13) & Chr(13) & "La longitud de la clave debera ser entre 6 a 8 caracteres." _
       & Chr(13) & "Esta última se cambiará pronto ampliando la longitud permitida.", vbInformation, sMessage
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
End Sub

Private Function fControlAlfa() As Boolean
   Dim mVect(3) As Boolean  '1-numerico / 2-alfa mayusc / 3-simbolo
   Dim mTecla As Integer
   Dim mI As Integer
   Dim mJ As Integer
   For mI = 1 To 3
      mVect(mI) = False
   Next
   mJ = 0
   For mI = 1 To Len(Trim(Text1(1).Text))
      mTecla = Asc(Mid(Trim(Text1(1).Text), mI, 1))
      If mTecla >= 48 And mTecla <= 57 Then 'numerico
         mVect(1) = True
         mJ = mJ + 1
      End If
      If (mTecla >= 65 And mTecla <= 90) Or (mTecla >= 97 And mTecla <= 122) Then  'mayusc
         mVect(2) = True
      End If
   Next
   If Not mVect(1) Or mJ < 2 Then
      MsgBox "La clave debe contener al menos dos número", vbInformation, sMessage
      mVect(1) = False
   End If
   If Not mVect(2) Then
      MsgBox "La clave debe contener al menos una letra", vbInformation, sMessage
   End If
   fControlAlfa = mVect(1) And mVect(2) 'And mVect(3)
End Function
