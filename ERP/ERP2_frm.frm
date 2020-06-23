VERSION 5.00
Begin VB.Form ERP2_frm 
   BackColor       =   &H00666666&
   Caption         =   "Sistema ERP - Permisos a Usuarios"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   Picture         =   "ERP2_frm.frx":0000
   ScaleHeight     =   7275
   ScaleWidth      =   9585
   Begin VB.CommandButton Command1 
      Caption         =   "Volver"
      Height          =   435
      Index           =   1
      Left            =   4260
      TabIndex        =   11
      Top             =   6600
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   0
      Left            =   4260
      TabIndex        =   10
      Top             =   5580
      Width           =   1155
   End
   Begin VB.ListBox List1 
      ForeColor       =   &H80000002&
      Height          =   3765
      Index           =   1
      Left            =   5520
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   3300
      Width           =   3975
   End
   Begin VB.ListBox List1 
      ForeColor       =   &H00404040&
      Height          =   3765
      Index           =   0
      Left            =   120
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   3300
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H8000000D&
      Height          =   315
      Index           =   2
      Left            =   3660
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2220
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H8000000D&
      Height          =   315
      Index           =   1
      Left            =   3660
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1800
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H8000000D&
      Height          =   315
      Index           =   0
      Left            =   3660
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1380
      Width           =   2655
   End
   Begin VB.Image Image3 
      Height          =   555
      Index           =   1
      Left            =   4560
      MouseIcon       =   "ERP2_frm.frx":14CF
      MousePointer    =   99  'Custom
      Picture         =   "ERP2_frm.frx":1621
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   555
   End
   Begin VB.Image Image3 
      Height          =   555
      Index           =   0
      Left            =   4560
      MouseIcon       =   "ERP2_frm.frx":1A0E
      MousePointer    =   99  'Custom
      Picture         =   "ERP2_frm.frx":1B60
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Permisos de usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   5580
      TabIndex        =   7
      Top             =   3120
      Width           =   1710
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menú del sistema"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   6
      Top             =   3120
      Width           =   1485
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   1380
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   2940
      TabIndex        =   4
      Top             =   2280
      Width           =   450
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   2760
      TabIndex        =   3
      Top             =   1860
      Width           =   660
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   1440
      Width           =   675
   End
End
Attribute VB_Name = "ERP2_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Me.Width = 9705
   Me.Height = 7680
   sAlinearForm Me
   sInitForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
   mExitSist 10
   ShowMenu 10, False, True
   ERP1_frm.Show
End Sub

Private Sub Combo1_Click(Index As Integer)
Dim mObj As New clLogUser
Dim mRec As New ADODB.Recordset
Dim mI As Integer

   Select Case Index
      Case 0, 1
         If Combo1(0).Text <> "" Then
            On Error Resume Next
            Image2.Picture = ERP1_frm.Image1(Trim(Right(Combo1(0).Text, 2))).Picture
         End If
         If Combo1(0).Text <> "" And Combo1(1).Text <> "" Then
            List1(0).Clear
            List1(1).Clear
            Set mRec = mObj.oMenuSistema(Trim(Right(Combo1(0).Text, 2)))
            Do While Not mRec.EOF
               List1(0).AddItem mRec!descripcion & Space(100) & mRec!CodigoMenu
               mRec.MoveNext
            Loop
            mRec.Close
            Set mRec = mObj.oPermisoMenu(Trim(Right(Combo1(1).Text, 15)), Trim(Right(Combo1(0).Text, 2)))
            Do While Not mRec.EOF
               List1(1).AddItem mRec!descripcion & Space(100) & mRec!codmenu
               For mI = 0 To List1(0).ListCount - 1
                  List1(0).ListIndex = mI
                  If Trim(Right(List1(0).Text, 25)) = Trim(mRec!codmenu) Then
                     List1(0).RemoveItem mI
                     mI = 999
                  End If
               Next
               mRec.MoveNext
            Loop
            mRec.Close
         End If
   
      Case 2
         Combo1(1).Clear
         If Combo1(2).ListIndex = 0 Then
            Set mRec = mObj.oTabla("usuarios", " where fechabaja is null order by apellido, nombres")
         Else
            Set mRec = mObj.oTabla("usuarios", " where fechabaja is null and codperfil='" & Right(Combo1(2).Text, 1) & "' order by  apellido, nombres")
         End If
         
         Do While Not mRec.EOF
            Combo1(1).AddItem mRec!apellido & ", " & mRec!nombres & Space(50) & mRec!CodUsuario
            mRec.MoveNext
         Loop
         mRec.Close
   End Select

Set mObj = Nothing
Set mRec = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mObj As New clLogUser
Dim mI As Integer

   If Index = 0 Then
      If Not mObj.xDeletePermisos(Trim(Right(Combo1(1).Text, 17)), Trim(Right(Combo1(0).Text, 2))) Then
         MsgBox "ERROR al borrar permisos...", vbCritical, sMessage
      End If
      For mI = 0 To List1(1).ListCount - 1
         List1(1).ListIndex = mI
         mObj.xInsertPermisos Trim(Right(Combo1(1).Text, 17)), Trim(Right(List1(1).Text, 30)), Trim(Right(Combo1(0).Text, 2))
            'MsgBox "ERROR de asignación...", vbCritical, sMessage
        ' End If
      Next
      Image2.Picture = LoadPicture()
   Else
      Unload Me
   End If
   Set mObj = Nothing
End Sub

Private Sub Image3_Click(Index As Integer)
Dim mI As Integer
Dim mJ As Integer
   For mI = 0 To List1(Index).ListCount - 1
      If List1(Index).Selected(mI) Then
         If Index = 0 Then
            List1(1).AddItem List1(0).List(mI)
         Else
            List1(0).AddItem List1(1).List(mI)
         End If
      End If
   Next
   mJ = 1
   mI = 0
   Do While mI < List1(Index).ListCount
      If List1(Index).Selected(mI) Then
         List1(Index).RemoveItem mI
      Else
         mI = mI + 1
      End If
   Loop
End Sub

Private Sub Image3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image3(Index).BorderStyle = 1
End Sub

Private Sub Image3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image3(Index).BorderStyle = 0
End Sub

Private Sub sInitForm()
Dim mObj As New clLogUser
Dim mRec As New ADODB.Recordset

   Set mRec = mObj.oTabla("usuarios", " WHERE FECHABAJA IS NULL ORDER BY apellido, nombres")
   Do While Not mRec.EOF
      Combo1(1).AddItem mRec!apellido & ", " & mRec!nombres & Space(50) & mRec!CodUsuario
      mRec.MoveNext
   Loop
   mRec.Close
   Set mRec = mObj.oTabla("sistemas", "WHERE FECHABAJA IS NULL ORDER BY DESCRIPCION")
   sLlenoCbo Combo1(0), mRec, 1, 0
   Combo1(2).AddItem "Todos" & Space(50) & "-1"
   Set mRec = mObj.oTabla("perfiles", "order by 1")
   sLlenoCbo Combo1(2), mRec, 1, 0
   
   Set mObj = Nothing
   Set mRec = Nothing
End Sub
