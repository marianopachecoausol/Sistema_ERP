VERSION 5.00
Begin VB.Form RNov1a_frm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   19605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   19605
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver pendientes"
      Height          =   375
      Left            =   14880
      TabIndex        =   13
      Top             =   240
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver todos"
      Height          =   375
      Left            =   13560
      TabIndex        =   12
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0EDEC&
      Caption         =   "depurar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Index           =   5
      Left            =   150
      Picture         =   "RNov1a_frm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   50
      Width           =   825
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C7A56B&
      Caption         =   "Rutinas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   2
      Left            =   5100
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   50
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Tareas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   1
      Left            =   2550
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   50
      Width           =   1200
   End
   Begin VB.ListBox List1 
      BackColor       =   &H8000000B&
      ForeColor       =   &H8000000D&
      Height          =   645
      ItemData        =   "RNov1a_frm.frx":03AE
      Left            =   10200
      List            =   "RNov1a_frm.frx":03B0
      TabIndex        =   6
      Top             =   180
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "Radio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   4
      Left            =   3825
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   50
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   850
      Index           =   1
      Left            =   13425
      TabIndex        =   4
      Top             =   0
      Width           =   60
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00B3C1CC&
      Caption         =   "Turnos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   3
      Left            =   1290
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   50
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C1DBD8&
      Caption         =   "Ingresar Novedad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   0
      Left            =   7425
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   25
      Width           =   2475
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   0
      Top             =   -150
   End
   Begin VB.Image Image4 
      Height          =   525
      Left            =   12600
      Picture         =   "RNov1a_frm.frx":03B2
      Stretch         =   -1  'True
      Tag             =   "nov.ico"
      Top             =   120
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "DEMORA PENDIENTE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   7425
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Image Image3 
      Height          =   705
      Left            =   11160
      Picture         =   "RNov1a_frm.frx":0CC4
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Móviles"
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
      Left            =   10200
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   5
      Left            =   17265
      Picture         =   "RNov1a_frm.frx":7FBE
      Stretch         =   -1  'True
      Top             =   135
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   4
      Left            =   16545
      Picture         =   "RNov1a_frm.frx":81BF
      Stretch         =   -1  'True
      Top             =   135
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   3
      Left            =   15825
      Picture         =   "RNov1a_frm.frx":83C0
      Stretch         =   -1  'True
      Top             =   135
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   2
      Left            =   15105
      Picture         =   "RNov1a_frm.frx":85C1
      Stretch         =   -1  'True
      Top             =   135
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   1
      Left            =   14385
      Picture         =   "RNov1a_frm.frx":87C2
      Stretch         =   -1  'True
      Top             =   135
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   0
      Left            =   13665
      Picture         =   "RNov1a_frm.frx":89C3
      Stretch         =   -1  'True
      Top             =   135
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   18135
      TabIndex        =   3
      Top             =   450
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   520
      Left            =   10
      Picture         =   "RNov1a_frm.frx":8BC4
      Stretch         =   -1  'True
      Tag             =   "nov.ico"
      Top             =   120
      Visible         =   0   'False
      Width           =   520
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Kilometraje"
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
      Height          =   255
      Left            =   18105
      TabIndex        =   2
      Top             =   180
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu MenuClima 
      Caption         =   "Clima"
      Visible         =   0   'False
      Begin VB.Menu MnuClima 
         Caption         =   "Despejado"
         Index           =   0
      End
      Begin VB.Menu MnuClima 
         Caption         =   "Lluvia"
         Index           =   1
      End
      Begin VB.Menu MnuClima 
         Caption         =   "Llovizna"
         Index           =   2
      End
      Begin VB.Menu MnuClima 
         Caption         =   "Lluvia Intensa"
         Index           =   3
      End
      Begin VB.Menu MnuClima 
         Caption         =   "Granizo"
         Index           =   4
      End
      Begin VB.Menu MnuClima 
         Caption         =   "Neblina"
         Index           =   5
      End
      Begin VB.Menu MnuClima 
         Caption         =   "Viento"
         Index           =   6
      End
      Begin VB.Menu MnuClima 
         Caption         =   "Nublado"
         Index           =   7
      End
   End
End
Attribute VB_Name = "RNov1a_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mRec As New ADODB.Recordset
Dim mRec1 As New ADODB.Recordset
Dim mImag As Integer
Dim mVerTodos As Boolean
Dim mPc As String

Private Sub Form_Load()
Dim mObj As New clRNov
Dim mi As Integer
Dim mj As Integer
  
   If Me.Option1.Value Then
      mVerTodos = True
   Else
      mVerTodos = False
   End If
   
   Me.Height = 960
   'Me.Width = 15300
   Me.Width = 19695
   Me.Top = 3170
   Me.Left = 0
   mPc = Mid(MDI.mPCname, 1, Len(MDI.mPCname) - 1)
   Set mRec = mObj.oTabla("actualizar", "WHERE flag=1 AND pc='" & Left(mPc, 12) & "'")
   If Not mRec.EOF Then
      mObj.xUpActualizar mPc, 0
      Image3.Visible = True
      PlaySound "ringin.wav"
   Else
      Image3.Visible = False
   End If
   mRec.Close
   If List1.ListCount > 0 Then
      List1.Visible = True
      Label3.Visible = True
   End If
   Set mRec = mObj.oTabla("climatag", "")
   If Not mRec.EOF Then
      For mi = 0 To 5
         mj = mRec!Tag
         sCargaClima Image2, mi, mj
         mRec.MoveNext
      Next
   End If
   mRec.Close
   Set mObj = Nothing
   
  ' MsgBox ("Top:" & Me.Top & "- Heigth:" & Me.Height)
End Sub

'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 '  Label1.Caption = ""
'End Sub

Private Sub Command2_Click(Index As Integer)
Dim mi As Integer
Dim mFlag As Boolean
   Select Case Index
      Case 0 'Ingreso de Novedades
         Command2(Index).Visible = False
         RNov1a_frm.Enabled = False
         RNov1d_frm.Enabled = False
         RNov1c_frm.Show
      Case 1, 2 'Tareas y Rutinas
         If List1.Visible Then
            mFlag = True
            For mi = 0 To List1.ListCount - 1
              If Left(List1.List(mi), 1) = "G" Then
                 mFlag = False
              End If
            Next
            If mFlag Then
               Command2(0).Visible = False
               RNov1c_frm.Show
               If Index = 1 Then
                  RNov1c_frm.sInitTareas  'Tareas
               Else
                  RNov1c_frm.sInitRutinas  'Rutinas
               End If
            Else
               MsgBox "Solo deberá existir Móviles Patrullas", vbCritical, sMessage
            End If
         Else
            MsgBox "Falta Asignar un Móvil", vbExclamation, sMessage
         End If
      Case 3 'Turnos de Móviles
         RNov5_frm.Show
         RNov1a_frm.Enabled = False
         RNov1b_frm.Enabled = False
         RNov1d_frm.Enabled = False
      Case 4 'Radio
         RNov3_frm.Show
         RNov3_frm.Frame1.Tag = "RADIO"
         RNov3_frm.sInitRadio
      Case 5 'depurar
         If MsgBox("Seguro de depurar la base de datos?", vbOKCancel, sMessage) = vbOK Then
            sDepurar
         End If
   End Select
End Sub

Private Sub Image1_Click()
   Image1.Visible = False
   Image3.Visible = False
   Unload RNov1b_frm
   Unload RNov1d_frm
   Load RNov1b_frm
   Load RNov1d_frm
End Sub

Private Sub Image3_Click()
Dim mObj As New clRNov
   Image3.Visible = False
   mObj.xUpActualizar mPc, 0
   RNov1a_frm.Label4.Visible = mObj.bExistDatoTabla("novedades2", "codnov='D'")
   Unload RNov1b_frm
   Unload RNov1d_frm
   Load RNov1b_frm
   Load RNov1d_frm
   'RNov1a_frm.Top = RNov1b_frm.Height + 20
   Set mObj = Nothing
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 1
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 0
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.BorderStyle = 1
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.BorderStyle = 0
End Sub

Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   mImag = Index
   PopupMenu MenuClima
End If
End Sub

Private Sub Image2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Select Case Index
      Case 0
          Label1.Caption = "(12.95 al 21.62)"
      Case 1
          Label1.Caption = "(21.62 al 35.84)"
      Case 2
          Label1.Caption = "(35.84 al 38.57)"
      Case 3
          Label1.Caption = "(38.57 al 47.66)"
      Case 4
          Label1.Caption = "(47.66 al 63.30)"
      Case 5
          Label1.Caption = "(63.30 al 65.14)"
   End Select
End Sub

Private Sub Image4_Click()
   RNov12.Show
   RNov1a_frm.Enabled = False
   RNov1b_frm.Enabled = False
   RNov1d_frm.Enabled = False
End Sub

Private Sub List1_DblClick()
   List1.RemoveItem List1.ListIndex
   If List1.ListCount < 1 Then
      List1.Visible = False
      Label3.Visible = False
   End If
End Sub

Private Sub MnuClima_Click(Index As Integer)
Dim mObj As New clRNov
   sCargaClima Image2, mImag, Index
   mObj.xUpClimaTag mImag, Index
   Set mObj = Nothing
End Sub

Private Sub sCargaClima(Imagen As Object, ByVal X As Integer, ByVal Indice As Integer)
   Select Case Indice
      Case 0
         Imagen(X).Picture = LoadPicture(App.Path & "\RegNovedades\Image\Despejado.gif")
      Case 1
         Imagen(X).Picture = LoadPicture(App.Path & "\RegNovedades\Image\Lluvia2.gif")
      Case 2
         Imagen(X).Picture = LoadPicture(App.Path & "\RegNovedades\Image\Llovisna.gif")
      Case 3
         Imagen(X).Picture = LoadPicture(App.Path & "\RegNovedades\Image\Tormenta.gif")
      Case 4
         Imagen(X).Picture = LoadPicture(App.Path & "\RegNovedades\Image\Granizo.gif")
      Case 5
         Imagen(X).Picture = LoadPicture(App.Path & "\RegNovedades\Image\Neblina.gif")
      Case 6
         Imagen(X).Picture = LoadPicture(App.Path & "\RegNovedades\Image\Viento.gif")
      Case 7
         Imagen(X).Picture = LoadPicture(App.Path & "\RegNovedades\Image\Nubes.gif")
   End Select
   Imagen(X).Tag = Indice
End Sub

Private Sub Option1_Click()
   mVerTodos = True
   Unload RNov1d_frm
   Load RNov1d_frm
End Sub

Private Sub Option2_Click()
   mVerTodos = False
   Unload RNov1d_frm
   Load RNov1d_frm
End Sub

Private Sub Timer1_Timer()
Dim mObj As New clRNov
Dim mRec2 As New ADODB.Recordset
   Set mRec2 = mObj.oTabla("actualizar", "where flag=1 AND PC='" & mPc & "'")
   If Not mRec2.EOF Then
      mObj.xUpActualizar mPc, 0
      Image3.Visible = True
   End If
   Set mRec2 = mObj.waze_getAccidentesLiberados
   If Not mRec2.EOF Then
      Image4.Visible = True
   Else
      Image4.Visible = False
   End If
   mRec2.Close
   Set mRec2 = Nothing
  If Image1.Visible Then
     PlaySound "LASER.WAV"
  End If
  If Image3.Visible Then
     PlaySound "ringin.wav"
  End If
  Set mObj = Nothing
End Sub

Private Sub sDepurar()
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
Dim mFlag As Boolean
Dim mFecha As String
Dim mKm As String

   mFecha = ""
   sMsgEspere Me, "Depurando... espere.", True
   Set mRec = mObj.oMaxTabla("novedades", "fecha", "")
   If Not mRec.EOF Then
      mFecha = mRec!Total
   End If
   mRec.Close
   Set mRec = mObj.oTabla("novedades2", "where fecha > '" & Format(mFecha, "yyyy-mm-dd hh:mm:ss") & "' order by fecha")
   If Not mRec.EOF Then
      Do While Not mRec.EOF
         mKm = mRec.Fields(3)
         mFlag = mObj.xInsNov_old(mRec.Fields(0), mRec.Fields(1), mRec.Fields(2), mKm, mRec.Fields(4), mRec.Fields(5), mRec.Fields(6), mRec.Fields(7), mRec.Fields(8), mRec.Fields(9), mRec.Fields(10), mRec.Fields(11), mRec.Fields(12), mRec.Fields(13), mRec.Fields(14), mRec.Fields(15), mRec.Fields(16), mRec.Fields(17), NVL(mRec.Fields(18), ""), mRec.Fields(19), mRec.Fields(20), mRec.Fields(21), mRec.Fields(22))
         mRec.MoveNext
      Loop
   End If
   mRec.Close
   mObj.xDepuraNov
   sMsgEspere Me, "", False
   MsgBox "Depuración Terminada!!!", vbInformation, sMessage
   
   Set mRec = Nothing
   Set mObj = Nothing
End Sub

Public Function verTodos() As Boolean
   verTodos = mVerTodos
End Function
