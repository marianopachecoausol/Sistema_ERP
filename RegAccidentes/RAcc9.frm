VERSION 5.00
Begin VB.Form RAcc9_frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulo de Comentarios"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   11775
   Begin VB.Frame Frame1 
      Height          =   7850
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   480
         TabIndex        =   9
         Top             =   240
         Width           =   3375
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Comentarios"
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
            Left            =   840
            TabIndex        =   10
            Top             =   200
            Width           =   1590
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   0
         Left            =   9840
         MaxLength       =   5
         TabIndex        =   8
         Text            =   "8888888"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   975
         Index           =   0
         Left            =   480
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1920
         Width           =   10695
      End
      Begin VB.TextBox Text2 
         Height          =   975
         Index           =   1
         Left            =   480
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   3240
         Width           =   10695
      End
      Begin VB.TextBox Text2 
         Height          =   975
         Index           =   2
         Left            =   480
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   4560
         Width           =   10695
      End
      Begin VB.TextBox Text2 
         Height          =   1215
         Index           =   3
         Left            =   480
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   5880
         Width           =   8295
      End
      Begin VB.Frame Frame13 
         Height          =   1695
         Left            =   9360
         TabIndex        =   1
         Top             =   5880
         Width           =   2055
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Grabar"
            Height          =   615
            Left            =   360
            Picture         =   "RAcc9.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Volver"
            Height          =   615
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "N° Orden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8520
         TabIndex        =   15
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descripción del Arribo al Lugar"
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
         Left            =   480
         TabIndex        =   14
         Top             =   1680
         Width           =   2640
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Comentario de los Involucrados o Testigos"
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
         Left            =   480
         TabIndex        =   13
         Top             =   3000
         Width           =   3630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mecánica del Accidente"
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
         Left            =   480
         TabIndex        =   12
         Top             =   4320
         Width           =   2070
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones del Patrullero"
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
         Left            =   480
         TabIndex        =   11
         Top             =   5640
         Width           =   2460
      End
   End
End
Attribute VB_Name = "RAcc9_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mRec As New ADODB.Recordset

Private Sub Form_Load()
   Me.Height = 8295
   Me.Width = 11900
   sAlinearForm Me
   Text1(0).Text = RAcc1_frm!Text2(2).Text
   If RAcc1_frm.mBusca Then
      Command2.Caption = "&Actualizar"
      Command2.BackColor = &H5256FE
      RAcc9_frm.Caption = RAcc9_frm.Caption & "  ---MODO VISTA Y MODIFICACIÓN--"
      Set mRec = RAcc1_frm.mObj.oTabla("fichadescr", "WHERE NroOrden='" & RAcc1_frm!Text2(2).Text & "'")
      If Not mRec.EOF Then
         Text2(0).Text = NVL(mRec!descr_arribo, "")
         Text2(1).Text = NVL(mRec!comentarios, "")
         Text2(2).Text = NVL(mRec!mecanica, "")
         Text2(3).Text = NVL(mRec!obs_patrullero, "")
      End If
   End If
End Sub

Private Sub Command1_Click()
   RAcc9_frm.Visible = False
   RAcc3_frm.Show
   RAcc3_frm.Top = 0
   RAcc3_frm.Left = 0
End Sub

Private Sub Command2_Click()
  Dim mI As Integer
  Dim mJ As Integer
  Dim OptionArray(14) As String
  Dim mArray(6) As String
  Dim mText As String
  Dim mFlag As Boolean
  
   If MsgBox("¿Desea Grabar los Datos?", vbYesNo, sMessage) = vbYes Then
      If Not RAcc1_frm.mBusca Then
         RAcc1_frm.mObj.xUpAuxiliar (Format((Val(RAcc1_frm!Text2(2).Text) + 1), "00000"))
      End If
       OptionArray(0) = SelectOption(RAcc1_frm!Option1) 'TRAMO
       mText = ""
      For mI = 0 To 6
         If RAcc1_frm!Check2(mI).Value = 1 Then
            mText = mText & Format((mI + 1), "00")
         End If
      Next
      OptionArray(1) = mText
      OptionArray(2) = SelectOption(RAcc1_frm!Option3) 'ACCIDE CON OTRO'
      OptionArray(3) = SelectOption(RAcc1_frm!Option4) 'OTRO
      OptionArray(4) = SelectOption(RAcc1_frm!Option5) 'SENTIDO TRANS
      OptionArray(5) = SelectOption(RAcc1_frm!Option6) 'LUGAR
      OptionArray(6) = SelectOption(RAcc1_frm!Option7) 'CLIMA
      OptionArray(7) = SelectOption(RAcc1_frm!Option9) 'CALZADA
      OptionArray(8) = SelectOption(RAcc1_frm!Option10) 'BANQUINA
      OptionArray(9) = SelectOption(RAcc1_frm!Option11) 'DEM. HOR
      OptionArray(10) = SelectOption(RAcc1_frm!Option12) 'DEM. VERT
      OptionArray(11) = SelectOption(RAcc1_frm!Option13) 'ILUM
      OptionArray(12) = SelectOption(RAcc1_frm!Option14) 'CAUSA VEHIC
      If RAcc3_frm!Check1(0).Value = 1 And RAcc3_frm!Check1(1).Value = 1 Then
         OptionArray(13) = "03"
      Else
         If RAcc3_frm!Check1(0).Value = 1 And RAcc3_frm!Check1(1).Value = 0 Then
            OptionArray(13) = "01"
         Else
            If RAcc3_frm!Check1(0).Value = 0 And RAcc3_frm!Check1(1).Value = 1 Then
               OptionArray(13) = "02"
            End If
         End If
      End If
              
      For mI = 0 To RAcc1_frm!Combo2.UBound
         If RAcc1_frm!Combo2(mI).Text <> "" Then
            mArray(mI) = Trim(Right(RAcc1_frm!Combo2(mI).Text, 2))
         End If
      Next
      mJ = 3
      For mI = 0 To RAcc1_frm!Combo3.UBound
         If RAcc1_frm!Combo3(mI).Text <> "" Then
            mArray(mJ) = Trim(Right(RAcc1_frm!Combo3(mI).Text, 3))
         End If
         mJ = mJ + 1
      Next
      If Not RAcc1_frm.mBusca Then
         MsgBox "Existe un módulo nuevo para el ingreso de Fichas de accidentes", vbInformation, sMessage
      Else
          mFlag = RAcc1_frm.mObj.xUpFichaOlder(RAcc1_frm!Text2(2).Text, Trim(Left(RAcc1_frm!Combo1.Text, 3)), Trim(RAcc1_frm!Text2(0).Text), Trim(RAcc1_frm!Text2(1).Text), Trim(RAcc1_frm!Text1(0).Text), Trim(RAcc1_frm!Text1(1).Text), Trim(RAcc1_frm!Text1(2).Text), Trim(RAcc1_frm!Text3(0).Text), OptionArray(0), RAcc1_frm!Text3(1).Text, OptionArray(1), RAcc1_frm!Text4.Text, OptionArray(2), mArray(0), mArray(1), mArray(2), OptionArray(3), OptionArray(4), _
                     OptionArray(5), OptionArray(6), OptionArray(7), OptionArray(8), OptionArray(9), OptionArray(10), OptionArray(11), mArray(3), mArray(4), mArray(5), OptionArray(12), Trim(RAcc3_frm!Text3(0).Text), Trim(RAcc3_frm!Text3(1).Text), _
                     Trim(RAcc3_frm!Text3(2).Text), Trim(RAcc3_frm!Text3(3).Text), Trim(RAcc3_frm!Text3(4).Text), Trim(RAcc3_frm!Text3(5).Text), Trim(RAcc3_frm!Text3(6).Text), Trim(RAcc3_frm!Text3(7).Text), Trim(RAcc3_frm!Text3(8).Text), Trim(RAcc3_frm!Text3(9).Text), Trim(RAcc3_frm!Text3(10).Text), Trim(RAcc3_frm!Text3(11).Text), Trim(RAcc3_frm!Text3(12).Text), Trim(RAcc3_frm!Text3(13).Text), OptionArray(13))
          
         RAcc1_frm.mObj.xUpdFichaDescr RAcc1_frm!Text2(2).Text, Trim(Text2(0).Text), Trim(Text2(1).Text), Trim(Text2(2).Text), Trim(Text2(3).Text)
         RAcc1_frm.mObj.xDelTabla "VehiculosInvolucr", "NroOrden = '" & Trim(RAcc1_frm!Text2(2).Text) & "'"
         RAcc1_frm.mObj.xDelTabla "VictimasInvolucr ", "NroOrden = '" & Trim(RAcc1_frm!Text2(2).Text) & "'"
      End If
      If RAcc2_frm!List1.ListCount > 0 Then
         For mI = 0 To RAcc2_frm!List1.ListCount - 1
            RAcc2_frm!List1.ListIndex = mI
            RAcc2_frm!List2.ListIndex = mI
            mFlag = RAcc1_frm.mObj.xInsVehicOlder(Trim(RAcc1_frm!Text2(2).Text), Trim(Mid(RAcc2_frm!List1.Text, 1, 2)), Trim(Mid(RAcc2_frm!List1.Text, 6, 2)), Trim(Mid(RAcc2_frm!List1.Text, 16, 2)), Trim(Mid(RAcc2_frm!List1.Text, 26, 15)), _
                  Trim(Mid(RAcc2_frm!List1.Text, 44, 8)), Trim(Mid(RAcc2_frm!List1.Text, 55, 25)), Trim(Mid(RAcc2_frm!List1.Text, 83, 2)), Trim(Mid(RAcc2_frm!List1.Text, 93, 9)), Trim(Mid(RAcc2_frm!List2.Text, 5, 50)), Trim(Mid(RAcc2_frm!List2.Text, 58, 15)), Trim(Mid(RAcc2_frm!List2.Text, 76, 2)), Trim(Mid(RAcc2_frm!List2.Text, 86, 20)))
         Next
      End If
      If RAcc3_frm.List1.ListCount > 0 Then
         For mI = 0 To RAcc3_frm.List1.ListCount - 1
            RAcc3_frm.List1.ListIndex = mI
            RAcc3_frm.List2.ListIndex = mI
            mFlag = RAcc1_frm.mObj.xInsVictimasOlder(RAcc3_frm!Text1(0).Text, Trim(Left(RAcc3_frm!List1.Text, 2)), Trim(Mid(RAcc3_frm!List1.Text, 6, 25)), Trim(Mid(RAcc3_frm!List1.Text, 34, 50)), Trim(Mid(RAcc3_frm!List1.Text, 87, 2)), Trim(Mid(RAcc3_frm!List1.Text, 97, 9)), Trim(Mid(RAcc3_frm!List2.Text, 6, 2)), Trim(Mid(RAcc3_frm!List2.Text, 16, 2)), Trim(Mid(RAcc3_frm!List2.Text, 31, 2)), Trim(Mid(RAcc3_frm!List2.Text, 48, 1)), Trim(Mid(RAcc3_frm!List2.Text, 55, 1)), Trim(Mid(RAcc3_frm!List2.Text, 60, 1)), Trim(Mid(RAcc3_frm!List2.Text, 75, 1)), Trim(Mid(RAcc3_frm!List2.Text, 68, 3)), Trim(Mid(RAcc3_frm!List2.Text, 80, 1)), Trim(Mid(RAcc3_frm!List2.Text, 91, 2)))
         Next
      End If
      Unload RAcc1_frm
      Unload RAcc2_frm
      Unload RAcc3_frm
      Unload RAcc9_frm
      RAcc1_frm.Show
      If Not RAcc1_frm.mBusca Then
         RAcc1_frm.Combo4.Visible = False
      End If
   End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
End Sub

Private Function SelectOption(mObj As Object) As String
Dim mI As Integer
   For mI = 0 To mObj.UBound
      If mObj(mI).Value Then
         SelectOption = Format((mI + 1), "00")
         mI = 60
      Else
         SelectOption = ""
      End If
   Next
End Function
