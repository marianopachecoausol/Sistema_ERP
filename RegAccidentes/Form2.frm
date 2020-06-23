VERSION 5.00
Begin VB.Form RAcc2_frm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Módulo de Vehículos Involucrados."
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11805
   ControlBox      =   0   'False
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7920
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   11535
      Begin VB.Frame Frame5 
         Height          =   1695
         Left            =   10200
         TabIndex        =   37
         Top             =   2160
         Width           =   1215
         Begin VB.CommandButton Command3 
            Caption         =   "Volver"
            Height          =   615
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   960
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Siguiente"
            Default         =   -1  'True
            Height          =   615
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   35
         Top             =   2280
         Width           =   9975
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   31
         Top             =   480
         Width           =   11055
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Ltr-Domicilio                                         - Teléfono        - CiaSeg    - N° Póliza"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   120
         TabIndex        =   36
         Top             =   2040
         Width           =   9975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Modelo          - Dominio  - Conductor                 - Tipo    - Nro Doc "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   2760
         TabIndex        =   34
         Top             =   240
         Width           =   7875
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Ltr - Tipo    - Marca   -"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   2625
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   5
         Left            =   4800
         MaxLength       =   15
         TabIndex        =   8
         Text            =   "Text10"
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Frame Frame4 
         Height          =   975
         Left            =   9720
         TabIndex        =   30
         Top             =   2640
         Width           =   1575
         Begin VB.CommandButton Command1 
            Caption         =   "Agregar"
            Height          =   615
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   360
         TabIndex        =   28
         Top             =   240
         Width           =   3855
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vehículos Involucrados"
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
            Left            =   360
            TabIndex        =   29
            Top             =   240
            Width           =   2940
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Height          =   555
         Index           =   1
         Left            =   10200
         MaxLength       =   2
         TabIndex        =   27
         Text            =   "AA"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   6
         Left            =   9120
         MaxLength       =   20
         TabIndex        =   11
         Text            =   "Text8"
         Top             =   2280
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   3
         Left            =   9120
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   4
         Left            =   4800
         MaxLength       =   50
         TabIndex        =   9
         Text            =   "Text7"
         Top             =   3120
         Width           =   4215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   3
         Left            =   6240
         MaxLength       =   9
         TabIndex        =   7
         Top             =   2160
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   2
         Left            =   4800
         MaxLength       =   25
         TabIndex        =   5
         Text            =   "Text4"
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   4
         Text            =   "Text3"
         Top             =   3260
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   2800
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2160
         Width           =   2655
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1680
         Width           =   2655
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
         Left            =   7800
         MaxLength       =   5
         TabIndex        =   24
         Text            =   "8888888"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Agregar"
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
         Left            =   1800
         TabIndex        =   41
         Top             =   2520
         Width           =   690
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Agregar"
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
         Left            =   10080
         TabIndex        =   40
         Top             =   2040
         Width           =   675
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono"
         Height          =   195
         Left            =   4080
         TabIndex        =   32
         Top             =   2640
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Letra"
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
         Left            =   9360
         TabIndex        =   26
         Top             =   720
         Width           =   645
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
         Left            =   6600
         TabIndex        =   25
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Póliza"
         Height          =   195
         Left            =   8280
         TabIndex        =   23
         Top             =   2280
         Width           =   765
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cía. de Seguro"
         Height          =   195
         Left            =   7920
         TabIndex        =   22
         Top             =   1680
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio"
         Height          =   195
         Left            =   4080
         TabIndex        =   21
         Top             =   3120
         Width           =   630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Nro."
         Height          =   195
         Left            =   5880
         TabIndex        =   20
         Top             =   2160
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Doc."
         Height          =   195
         Left            =   3960
         TabIndex        =   19
         Top             =   2160
         Width           =   705
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Conductor"
         Height          =   195
         Left            =   3960
         TabIndex        =   18
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dominio"
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   3280
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Modelo"
         Height          =   195
         Left            =   360
         TabIndex        =   16
         Top             =   2830
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   2160
         Width           =   450
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Tipo "
         Height          =   195
         Left            =   480
         TabIndex        =   14
         Top             =   1680
         Width           =   360
      End
   End
End
Attribute VB_Name = "RAcc2_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mRec As New ADODB.Recordset
Dim mLetra As String
Public mAgregar As Boolean
Dim mI As Integer

Private Sub Form_Load()
   sInitForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mRec = Nothing
End Sub

Private Sub Combo1_Click(Index As Integer)
Dim mCodigo As String

   If Index = 0 Then
      mCodigo = mId(Combo1(0).Text, 1, 2)
      If mCodigo <> "12" And mCodigo <> "13" And mCodigo <> "14" Then
         Combo1(1).Enabled = True
         Combo1(1).Clear
         Set mRec = RAcc1_frm.mObj.oTablaCodigo("Marca", " codtipovehic='" & mCodigo & "' order by descripcion")
         sLlenoCbo RAcc2_frm.Combo1(1), mRec, 2, 1
       Else
          Combo1(1).Clear
          Combo1(1).Enabled = False
       End If
   End If
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 27 Then
      Combo1(Index).ListIndex = -1
   End If
End Sub

Private Sub Command1_Click()
  RAcc1_frm.Visible = False
  Dim mChar1, mChar2 As Integer
  If Command1.Caption = "Actualizar" Then
     If fLlenaLista Then
        sLimpiarForm2
        Text1(1).Text = mLetra
        Command1.Caption = "Otro"
        Command1.Picture = LoadPicture("checkmrk.ico")
        MsgBox "Registro Actualizado", vbInformation, sMessage
     End If
  Else
    If fLlenaLista Then
       If Len(Text1(1).Text) = 2 Then
          mChar1 = Asc(mId(Text1(1).Text, 2, 1))
          mChar2 = Asc(mId(Text1(1).Text, 1, 1))
          If mChar1 = 90 Then
             If mChar2 <> 90 Then
               Text1(1).Text = "" & Chr(mChar2 + 1) & "A"
             End If
          Else
             Text1(1).Text = "" & Chr(mChar2) & "" & Chr(mChar1 + 1) & ""
          End If
       Else
          mChar1 = Asc(Text1(1).Text)
          If mChar1 = 90 Then
            Text1(1).Text = "AA"
          Else
            Text1(1).Text = "" & Chr(mChar1 + 1) & ""
          End If
       End If
       sLimpiarForm2
    End If
  End If
End Sub

Private Sub Command2_Click()
  Dim mI As Integer
  If fTodoVacio Then
     RAcc1_frm.Visible = False
     RAcc2_frm.Visible = False
     RAcc3_frm!Combo1(0).Clear
     RAcc3_frm!Combo1(5).Clear
     If List1.ListCount <> 0 Then
       For mI = 0 To List1.ListCount - 1
         List1.ListIndex = mI
         RAcc3_frm!Combo1(0).AddItem Trim(mId(List1.Text, 55, 25)), mI  'Nombre Conductor
         RAcc3_frm!Combo1(5).AddItem Trim(mId(List1.Text, 1, 2)), mI  'Letra
       Next
       End If
     RAcc3_frm.Visible = True
     RAcc3_frm.Top = 0
     RAcc3_frm.Left = 0
  Else
     MsgBox "Tiene Datos Sin Agregar", vbCritical, sMessage
  End If
  If RAcc1_frm.mBusca Then
     RAcc3_frm.Text1(0).Text = RAcc1_frm.mNroOrden
  End If
End Sub

Private Sub Command3_Click()
  Me.Visible = False
  RAcc1_frm.Show
  RAcc1_frm.Top = 0
  RAcc1_frm.Left = 0
End Sub

Private Sub Label10_Click()
   If Label10.Caption = "Agregar" Then
      RAcc4_frm.mTabla = "CiaSeguros"
      RAcc4_frm.Label1.Caption = "Tabla de Cías de Seguros"
      RAcc4_frm.Show
      Label10.Caption = "Actualizar"
      Combo1(3).Clear
   Else
      Set mRec = RAcc1_frm.mObj.oTabla("CiaSeguros", "order by descripcion")
      sLlenoCbo Me.Combo1(3), mRec, 1, 0
      Label10.Caption = "Agregar"
   End If
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Label10.BorderStyle = 1
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Label10.BorderStyle = 0
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Label19.BorderStyle = 1
End Sub

Private Sub Label19_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Label19.BorderStyle = 0
   If Label19.Caption = "Agregar" Then
      mAgregar = True
      RAcc5_frm.Show
      Label19.Caption = "Actualizar"
      Combo1(1).Clear
   Else
      Combo1(0).ListIndex = -1
      Combo1(1).Clear
      Label19.Caption = "Agregar"
   End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii >= 97 And KeyAscii <= 122 Then
      KeyAscii = KeyAscii - 32
   End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 2
           KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
      Case 3
           KeyAscii = fNumeroKeyPress(KeyAscii)
      Case 5
           If KeyAscii <> 45 Then
             KeyAscii = fNumeroKeyPress(KeyAscii)
           End If
      Case 6
            KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
      Case Else
           KeyAscii = fAlfaNumKeyPress(KeyAscii)
   End Select
End Sub

Private Sub List1_Click()
   List2.ListIndex = List1.ListIndex
End Sub

Private Sub List2_Click()
   List1.ListIndex = List2.ListIndex
End Sub

Private Sub List1_DblClick()
Dim mA As String
Dim mB As String
Dim mResp As Integer
Dim mJ As Integer
Dim mFlag As Boolean
   mLetra = Text1(1).Text
   mFlag = True
   List2.ListIndex = List1.ListIndex
   If (List1.ListCount - 1) = List1.ListIndex Then
      If MsgBox("¿Desea Eliminar el Registro?", vbYesNo, "Atención!!!", sMessage) = vbYes Then
         For mI = 0 To RAcc3_frm!List2.ListCount - 1
            RAcc3_frm!List2.ListIndex = mI
            mA = Trim(mId(List1.Text, 1, 2))
            mB = Trim(mId(RAcc3_frm!List2.Text, 60, 2))
            If mA = mB Then
                mJ = 20
            End If
         Next
         If mJ = 20 Then
            MsgBox "Tiene Datos Cargados en Victimas con Vehículo -" & Trim(mId(List1.Text, 1, 2)) & "", vbCritical, sMessage
            mFlag = False
         Else
            mJ = List1.ListIndex
            Text1(1).Text = Right(mId(List1.Text, 1, 2), 2)
            List1.RemoveItem (mJ)
            List2.RemoveItem (mJ)
            mFlag = False
            RAcc2_frm.SetFocus
         End If
      End If
   End If
   If mFlag Then
      If MsgBox("¿Desea Actualizar el Registro?", vbYesNo, "Atención!!!", sMessage) = vbYes Then
         RAcc2_frm.SetFocus
         mLetra = Text1(1).Text
         Text1(1).Text = Trim(mId(List1.Text, 1, 2))            'Llena Letra
         mI = Trim(mId(List1.Text, 5, 3))
         Combo1(0).ListIndex = mI - 1             'Llena Tipo Vehículo
         If Trim(mId(List1.Text, 16, 2)) <> "" Then
            mJ = Trim(mId(List1.Text, 16, 2))
            For mI = 0 To Combo1(1).ListCount - 1                           'Llena Marca
               If mJ = Trim(Right(Combo1(1).List(mI), 2)) Then
                  Combo1(1).ListIndex = mI
               End If
            Next
         End If
         Text2(0) = Trim(mId(List1.Text, 26, 15))     'Llena Modelo
         Text2(1) = Trim(mId(List1.Text, 44, 8))     'Llena Dominio
         Text2(2) = Trim(mId(List1.Text, 55, 25))    'Llena Conductor
         If Trim(mId(List1.Text, 83, 2)) <> "" Then
            mJ = Right(mId(List1.Text, 83, 2), 2)
            For mI = 0 To Combo1(2).ListCount - 1
               Combo1(2).ListIndex = mI            'Llena Tipo Docu
               If mJ = Right(Combo1(2).Text, 3) Then
                  mI = 50
               End If
            Next
         End If
         Text2(3).Text = Trim(mId(List1.Text, 92, 9))  'Llena Nro Docu
      'List 2 **************
         Text2(4).Text = Trim(mId(List2.Text, 5, 50))   'Llena Dirección
         Text2(5).Text = Trim(mId(List2.Text, 58, 15))  'Llena Teléfono
         If Trim(mId(List2.Text, 76, 3)) <> "" Then
            mJ = Right(mId(List2.Text, 76, 3), 3)
            For mI = 0 To Combo1(3).ListCount - 1
               Combo1(3).ListIndex = mI            ' Llena Cía Seguro
               If mJ = Right(Combo1(3).Text, 3) Then
                  mI = 500
               End If
            Next
         End If
         Text2(6).Text = Trim(mId(List2.Text, 86, 20))
         
         Command1.Caption = "Actualizar"
         Command1.Picture = LoadPicture("erase02.ico")
         mJ = List1.ListIndex
         List1.RemoveItem (mJ)
         List2.RemoveItem (mJ)
       End If
    End If
End Sub

Private Sub List2_DblClick()
    List1.ListIndex = List2.ListIndex
    List1_DblClick
End Sub

Private Sub sLimpiarForm2()
   For mI = 0 To Combo1.UBound
     Combo1(mI).ListIndex = -1
   Next
   Combo1(1).Clear
   Combo1(1).Enabled = False
   For mI = 0 To Text2.UBound
       Text2(mI).Text = ""
   Next
End Sub

Private Function fLlenaLista() As Boolean
   fLlenaLista = True
   If Combo1(0).ListIndex <> -1 Then
                        '********LETRA************************  /  ***********************TIPO DE VEHICULO ********************************************   /   ************************ MARCA ***********************************************     /    ************* MODELO  ***************   /    ********** DOMINIO  ******************   /   ***********  CONDUCTOR    *****************
      List1.AddItem "" & Right(Space(2) & Text1(1).Text, 2) & " - " & Right(Space(7) & mId(Combo1(0).Text, 1, 2) & " " & mId(Combo1(0).Text, 7, 4), 7) & " - " & Right(Space(2) & Combo1(1).Text, 2) & " " & Left(Combo1(1).Text & Space(4), 4) & " - " & Left(Text2(0).Text & Space(15), 15) & " - " & Left(Text2(1).Text & Space(8), 8) & " - " & Left(Text2(2).Text & Space(25), 25) & " - " _
                  & "" & Right(Space(7) & (Right(Combo1(2).Text, 2)) & " " & Left(Combo1(2).Text, 4), 7) & " - " & Left(Text2(3).Text & Space(10), 10) & ""
                       '***************** TIPO  DOCU  ****************************************************   /  ***
                       
                       '************  LETRA  **************    /    *************  DIRECCION  *************  /    **********  TELEFONO   ****************   /  ******** TIPO DOCU ********************    /
      List2.AddItem "" & Right(Space(2) & Text1(1).Text, 2) & "- " & Left(Text2(4).Text & Space(50), 50) & " - " & Left(Text2(5).Text & Space(15), 15) & " - " & Right(Space(3) & Combo1(3).Text, 2) & " " & Left(Combo1(3).Text & Space(4), 4) & " - " & Left(Text2(6).Text & Space(20), 20) & ""
   Else
      MsgBox "Debe Seleccionar un tipo de Vehículo", vbCritical, "Atención!!!", sMessage
      fLlenaLista = False
   End If
End Function

Private Function fTodoVacio() As Boolean
Dim mI As Integer
fTodoVacio = True
For mI = 0 To 3
  fTodoVacio = fTodoVacio And Trim(Combo1(mI).Text) = ""
Next
For mI = 0 To 6
  fTodoVacio = fTodoVacio And Trim(Text2(mI).Text) = ""
Next
End Function

Private Sub sInitForm()
Dim mRS1 As New ADODB.Recordset
Dim mTipoVehic As String
Dim mMarca As String
Dim mTipoDoc As String
Dim mSeguro As String
Dim mChar1 As Integer
Dim mChar2 As Integer
Dim xLetra As String
  
   Me.Height = 8300
   Me.Width = 11900
   Me.Top = 0
   Me.Left = 0
   Unload RAcc3_frm
   mAgregar = False
   Combo1(1).Enabled = False
   Text1(0).Text = RAcc1_frm!Text2(2).Text
   Text1(1).Text = Chr(65)
   sLimpiarForm2
   Set mRec = RAcc1_frm.mObj.oTabla("TipoVehiculo", "order by codtipovehic")
   Do While Not mRec.EOF
      Combo1(0).AddItem "" & mRec!CodTipoVehic & " -  " & mRec!Descripcion & ""
      mRec.MoveNext
   Loop
   mRec.Close
   Set mRec = RAcc1_frm.mObj.oTabla("TipoDocu", "")
   sLlenoCbo Me.Combo1(0), mRec, 1, 0
   Set mRec = RAcc1_frm.mObj.oTabla("CiaSeguros", "order by descripcion")
   sLlenoCbo Me.Combo1(3), mRec, 1, 0
   If RAcc1_frm.mBusca Then
      Me.Caption = Me.Caption & "  ---MODO VISTA Y MODIFICACIÓN--"
      xLetra = ""
      Set mRec = RAcc1_frm.mObj.oTabla("VehiculosInvolucr", " where nroorden='" & RAcc1_frm!Text2(2).Text & "'")
      Do While Not mRec.EOF
         mMarca = ""
         mTipoDoc = ""
         mSeguro = ""
         mTipoVehic = Left(RAcc1_frm.mObj.sTablaDescr("TipoVehiculo", "Codtipovehic='" & mRec!CodTipoVehic & "'", 1), 4)
         If mRec!CodMarca <> "" Then
            mMarca = RAcc1_frm.mObj.sTablaDescr("Marca", "codtipovehic='" & mRec!CodTipoVehic & "' and codmarca='" & mRec!CodMarca & "'", 2)
         End If
         If mRec!codtipodoc <> "" Then
            mTipoDoc = RAcc1_frm.mObj.sTablaDescr("TipoDocu", "codtipodocu='" & mRec!codtipodoc & "'", 1)
         End If
         If mRec!CodCiaSeguro <> "" Then
            mSeguro = RAcc1_frm.mObj.sTablaDescr("CiaSeguros", "codciaseguro='" & mRec!CodCiaSeguro & "'", 1)
         End If
                          '************  LETRA  **************    /    *************  DIRECCION  *************  /    **********  TELEFONO   ****************   /  ******** TIPO DOCU ********************    /
         List1.AddItem "" & Right(Space(2) & mRec!letra, 2) & " - " & Right(Space(7) & mRec!CodTipoVehic & " " & mTipoVehic, 7) & " - " & Right(Space(7) & mRec!CodMarca & " " & mMarca, 7) & " - " & Left(mRec!modelo & Space(15), 15) & " - " & Left(mRec!Dominio & Space(8), 8) & " - " & Left(mRec!conductor & Space(25), 25) & " - " & Left(mRec!codtipodoc & " " & mTipoDoc & Space(7), 7) & " - " & Left(Trim(mRec!nrodocu & Space(10)), 10) & ""
         List2.AddItem "" & Right(Space(2) & mRec!letra, 2) & "- " & Left(Trim(mRec!domicilio) & Space(50), 50) & " - " & Left(mRec!Telefono & Space(15), 15) & " - " & Right(Space(7) & mRec!CodCiaSeguro & " " & mSeguro, 7) & " - " & Left(mRec!NroPoliza & Space(20), 20) & ""
         xLetra = mRec!letra
         mRec.MoveNext
      Loop
      mRec.Close
      If xLetra <> "" Then
         Text1(1).Text = xLetra
         If Len(Text1(1).Text) = 2 Then
            mChar1 = Asc(mId(Text1(1).Text, 2, 1))
            mChar2 = Asc(mId(Text1(1).Text, 1, 1))
            If mChar1 = 90 Then
               If mChar2 <> 90 Then
                  Text1(1).Text = "" & Chr(mChar2 + 1) & "A"
               End If
            Else
               Text1(1).Text = "" & Chr(mChar2) & "" & Chr(mChar1 + 1) & ""
            End If
         Else
             mChar1 = Asc(Text1(1).Text)
             If mChar1 = 90 Then
                Text1(1).Text = "AA"
             Else
                Text1(1).Text = "" & Chr(mChar1 + 1) & ""
            End If
         End If
      Else
         Text1(1).Text = "A"
      End If
   End If
   Set mRS1 = Nothing
End Sub
