VERSION 5.00
Begin VB.Form ERP5_frm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ERP - Módulo de ABM de Usuarios"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   8070
   Begin VB.Frame Frame2 
      Height          =   3975
      Left            =   60
      TabIndex        =   12
      Top             =   840
      Width           =   6375
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2040
         MaxLength       =   35
         TabIndex        =   3
         Top             =   2100
         Width           =   3255
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   4
         Top             =   2600
         Width           =   3255
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2040
         MaxLength       =   35
         TabIndex        =   2
         Top             =   1620
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   2040
         MaxLength       =   13
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nombres"
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
         Index           =   4
         Left            =   1200
         TabIndex        =   28
         Top             =   2160
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Desbloquear Usuario"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   4500
         TabIndex        =   27
         Top             =   1260
         Width           =   1500
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   1
         Left            =   4020
         MouseIcon       =   "ERP5_frm.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ERP5_frm.frx":0152
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Blanquear Clave"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   2520
         TabIndex        =   26
         Top             =   1260
         Width           =   1170
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   0
         Left            =   2040
         MouseIcon       =   "ERP5_frm.frx":0B54
         MousePointer    =   99  'Custom
         Picture         =   "ERP5_frm.frx":0CA6
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código"
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
         Index           =   3
         Left            =   1260
         TabIndex        =   18
         Top             =   790
         Width           =   600
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "E-Mail"
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
         TabIndex        =   17
         Top             =   2660
         Width           =   540
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Perfil"
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
         Left            =   1440
         TabIndex        =   16
         Top             =   3180
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Apellido"
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
         Left            =   1200
         TabIndex        =   15
         Top             =   1680
         Width           =   690
      End
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Index           =   1
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   2400
      Width           =   2055
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Index           =   0
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Height          =   3975
      Left            =   6480
      TabIndex        =   13
      Top             =   860
      Width           =   1575
      Begin VB.CommandButton Command1 
         Caption         =   "Clonar"
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
         Index           =   4
         Left            =   120
         MouseIcon       =   "ERP5_frm.frx":2D18
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Baja"
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
         Index           =   2
         Left            =   120
         MouseIcon       =   "ERP5_frm.frx":2E6A
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Modificar"
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
         Left            =   120
         MouseIcon       =   "ERP5_frm.frx":2FBC
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Alta"
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
         Left            =   120
         MouseIcon       =   "ERP5_frm.frx":310E
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Volver"
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
         Index           =   3
         Left            =   120
         MouseIcon       =   "ERP5_frm.frx":3260
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   3360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7990
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   4560
         TabIndex        =   19
         Top             =   360
         Width           =   3075
      End
      Begin VB.Image Image1 
         Height          =   660
         Left            =   240
         Picture         =   "ERP5_frm.frx":33B2
         Stretch         =   -1  'True
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ABM de Usuarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   1200
         TabIndex        =   14
         Top             =   240
         Width           =   2190
      End
   End
   Begin VB.Label Label5 
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
      Left            =   3240
      TabIndex        =   25
      Top             =   2400
      Width           =   3075
   End
   Begin VB.Label Label5 
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
      Left            =   3240
      TabIndex        =   24
      Top             =   1800
      Width           =   3075
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Destino"
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
      Left            =   240
      TabIndex        =   23
      Top             =   2445
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Origen"
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
      Left            =   360
      TabIndex        =   21
      Top             =   1845
      Width           =   570
   End
End
Attribute VB_Name = "ERP5_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Me.Height = 5265
   Me.Width = 8190
   sAlinearForm Me
   sInitForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ShowMenu 11, False, True
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mFlag As Boolean
Dim mResp As String
Dim mi As Integer
   Select Case Index
      Case 0
         If Command1(0).Caption = "Grabar" Then
            If fValid() = True Then
               sGrabar
               sInitForm
            End If
         Else
            Command1(0).Caption = "Grabar"
            Command1(1).Caption = "Cancelar"
            Command1(2).Visible = False
            Command1(4).Visible = False
            Combo2.Visible = False
            Label2.Caption = "Alta"
            Text1(0).Visible = True
            For mi = 0 To Text1.UBound
               Text1(mi).Text = ""
               Text1(mi).Enabled = True
            Next
            Combo1.ListIndex = -1
            Combo1.Enabled = True
            sIconos False
         End If
         
      Case 1
         If Command1(1).Caption = "Cancelar" Then
             sInitForm
          Else
             Command1(0).Caption = "Grabar"
             Command1(1).Caption = "Cancelar"
             Command1(2).Visible = False
             Combo2.Enabled = False
             Text1(1).Enabled = True
             Text1(2).Enabled = True
             Text1(3).Enabled = True
             Combo1.Enabled = True
             Label2.Caption = "Modificar"
             sIconos False
          End If
      
      Case 2
         If Combo1.ListIndex > -1 Then
            If MsgBox("Está Seguro de Eliminar al Usuario " & Text1(1).Text & "", vbOKCancel, sMessage) = vbOK Then
               sEliminar
            End If
         Else
            MsgBox "Seleccionar un usuario.", vbInformation, sMessage
         End If
         
      Case 3
          mExitSist 11
          Unload ERP5_frm
          ERP1_frm.Show
          
      Case 4 'Clonar Usuario
          If Command1(4).Caption = "Clonar" Then
              Frame2.Visible = False
              Command1(4).Caption = "Grabar"
              Command1(1).Caption = "Cancelar"
              Command1(0).Visible = False
              Command1(2).Visible = False
              Command1(3).Visible = False
              Combo3(0).ListIndex = -1
              Combo3(1).ListIndex = -1
              sIconos False
           Else
              If MsgBox("Está Seguro de Clonar el Usuario " & Combo3(0).Text & "", vbOKCancel, sMessage) = vbOK Then
                 sClonar
              End If
         End If
   End Select
End Sub

Private Sub Combo2_Click()
Dim mi As Integer
   If Combo2.Text <> "" Then
      Image2(0).Visible = True
      Image2(1).Visible = True
      Label6(0).Visible = True
      Label6(1).Visible = True
      sDatosUser
   Else
      Image2(0).Visible = False
      Image2(1).Visible = False
      Label6(0).Visible = False
      Label6(1).Visible = False
   End If
   
End Sub

Private Sub Combo3_Click(Index As Integer)
   If Trim(Combo3(Index).Text) <> "" Then    '
      sDatoName Index
   Else
      Label5(Index).Caption = ""
   End If
End Sub

Private Sub Image2_Click(Index As Integer)
Dim mObj As New clLogUser
   If Index = 0 Then
      If MsgBox("Está Seguro de Blanquear la clave del Usuario " & Combo2.Text & "", vbOKCancel, sMessage) = vbOK Then
         If mObj.xUpResetClaveUser(Combo2.Text) Then
            MsgBox "Clave Reseteada con Éxito", vbInformation, sMessage
         End If
      End If
   Else
      If MsgBox("Está Seguro de desbloquear al Usuario " & Combo2.Text & "", vbOKCancel, sMessage) = vbOK Then
         If mObj.xUpUnLockUser(Combo2.Text) Then
            MsgBox "Operación exitosa.", vbInformation, sMessage
         End If
      End If
   End If
   Set mObj = Nothing
End Sub

Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image2(Index).BorderStyle = 1
End Sub

Private Sub Image2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image2(Index).BorderStyle = 0
End Sub

'------------------------------------------------------------------------------------------------------
' FUNCIONES Y PROCESOS
'------------------------------------------------------------------------------------------------------

Private Function sInitForm()
Dim mObj As New clLogUser
Dim mRec As New ADODB.Recordset

   Frame2.Visible = True
   Combo1.Clear
   Combo2.Clear
   Combo2.Visible = True
   Combo2.Enabled = True
   Combo3(0).Clear
   Combo3(1).Clear
   Command1(0).Caption = "Alta"
   Command1(1).Caption = "Modificar"
   Command1(2).Caption = "Baja"
   Command1(3).Caption = "Volver"
   Command1(4).Caption = "Clonar"
   Command1(0).Visible = True
   Command1(1).Visible = True
   Command1(2).Visible = True
   Command1(3).Visible = True
   Command1(4).Visible = True
   Label2.Caption = ""
   sIconos False
   Text1(0).Text = ""
   Text1(0).Visible = False
   Text1(1).Text = ""
   Text1(2).Text = ""
   Text1(3).Text = ""
   Set mRec = mObj.oTabla("usuarios", " WHERE FechaBaja IS NULL ORDER BY CodUsuario")
   Do While Not mRec.EOF
      Combo2.AddItem mRec!CodUsuario
      Combo3(0).AddItem mRec!CodUsuario
      Combo3(1).AddItem mRec!CodUsuario
      mRec.MoveNext
   Loop
   mRec.Close
   Combo2.ListIndex = -1
   
   Set mRec = mObj.oTabla("perfiles", "order by 1")
   sLlenoCbo Combo1, mRec, 1, 0
   
'   Combo3(0).Clear
'   Combo3(1).Clear
   Label5(0).Caption = ""
   Label5(1).Caption = ""
   
   Set mObj = Nothing
   Set mRec = Nothing
End Function

Private Sub sIconos(ByVal pFlag As Boolean)
   Image2(0).Visible = pFlag
   Image2(1).Visible = pFlag
   Label6(0).Visible = pFlag
   Label6(1).Visible = pFlag
End Sub

Private Sub sGrabar()
Dim mObj As New clLogUser

   If Combo2.Visible Then 'Es Modificar
      mObj.xUpDatosUser Trim(Combo2.Text), Trim(Right(Combo1.Text, 2)), Trim(Text1(1).Text), Trim(Text1(2).Text), Trim(Text1(3).Text)
   Else 'Es Alta
      mObj.xInsNewUser Trim(LCase(Text1(0).Text)), Trim(Text1(1).Text), Trim(Text1(2).Text), Trim(Right(Combo1.Text, 2)), Trim(Text1(3).Text)
   End If
   Set mObj = Nothing
End Sub

Private Function fValid() As Boolean
Dim mObj As New clLogUser
Dim mRec As New ADODB.Recordset
Dim mText As String
   
   mText = ""
   fValid = True
   If Label2.Caption = "Alta" Then
      If Trim(Text1(0).Text) = "" Then mText = "_ Código de usuario. " & Chr(13)
   End If
   If Trim(Text1(1).Text) = "" Then mText = mText & "_ Apellido. " & Chr(13)
   If Trim(Text1(2).Text) = "" Then mText = mText & "_ Nombres. " & Chr(13)
   Set mRec = mObj.oTabla("usuarios", " where codusuario='" & Trim(Text1(0).Text) & "'")
   If Not mRec.EOF Then mText = mText & "_ USUARIO EXISTENTE. " & Chr(13)
   mRec.Close
   If mText <> "" Then
      MsgBox "Verificar los siguientes campos: " & Chr(13) & Chr(13) & mText
      fValid = False
   End If
   Set mObj = Nothing
   Set mRec = Nothing
End Function

Private Sub sEliminar()
Dim mObj As New clLogUser
Dim mRec As New ADODB.Recordset
Dim mi As Integer
   
   mObj.xUpBajaUser Trim(Combo2.Text)
   Combo2.Clear
   Combo3(0).Clear
   Combo3(1).Clear
   For mi = 0 To Text1.UBound
      Text1(mi) = ""
   Next
   Combo1.ListIndex = -1
   Set mRec = mObj.oTabla("usuarios", " WHERE FechaBaja IS NULL ORDER BY CodUsuario")
   Do While Not mRec.EOF
      Combo2.AddItem mRec!CodUsuario
      Combo3(0).AddItem mRec!CodUsuario
      Combo3(1).AddItem mRec!CodUsuario
      mRec.MoveNext
   Loop
   mRec.Close
   Combo2.ListIndex = -1
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Private Sub sClonar()
Dim mObj As New clLogUser
Dim mRec As New ADODB.Recordset

   Set mRec = mObj.oTabla("permisos", " where codusuario='" & Combo3(0).Text & "'")
   Do While Not mRec.EOF
      mObj.xInsertPermisos Combo3(1).Text, mRec!codmenu, mRec!codsistema
      mRec.MoveNext
   Loop
   mRec.Close
   MsgBox "Usuario Clonado", vbInformation, sMessage
   sInitForm
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Private Sub sDatosUser()
Dim mObj As New clLogUser
Dim mRec As New ADODB.Recordset
Dim mi As Integer
   
   Set mRec = mObj.oTabla("usuarios", "WHERE CodUsuario='" & Trim(Combo2.Text) & "'")
   If Not mRec.EOF Then
      Text1(1).Text = mRec!apellido
      Text1(2).Text = mRec!nombres
      Text1(3).Text = mRec!Email
      For mi = 0 To Combo1.ListCount - 1
         If CInt(Trim(Right(Combo1.List(mi), 2))) = mRec!CodPerfil Then
            Combo1.ListIndex = mi
         End If
      Next
   End If
   mRec.Close
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Private Sub sDatoName(ByVal pIndex As Integer)
Dim mObj As New clLogUser
Dim mRec As New ADODB.Recordset

   Set mRec = mObj.oTabla("usuarios", " WHERE CodUsuario='" & Trim(Combo3(pIndex).Text) & "'")
   If Not mRec.EOF Then
      Label5(pIndex).Caption = mRec!apellido & ", " & mRec!nombres
   End If
   mRec.Close
      
   Set mObj = Nothing
   Set mRec = Nothing
End Sub
