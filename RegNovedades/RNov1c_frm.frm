VERSION 5.00
Begin VB.Form RNov1c_frm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ingreso de Novedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   5035
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   9400
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   5
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2345
         Width           =   4700
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   4
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   915
         Left            =   75
         TabIndex        =   32
         Top             =   3950
         Visible         =   0   'False
         Width           =   6510
         Begin VB.CheckBox Check2 
            Height          =   240
            Index           =   10
            Left            =   6150
            TabIndex        =   18
            Top             =   375
            Width           =   240
         End
         Begin VB.CheckBox Check2 
            Height          =   240
            Index           =   9
            Left            =   5650
            TabIndex        =   17
            Top             =   375
            Width           =   240
         End
         Begin VB.CheckBox Check2 
            Height          =   240
            Index           =   8
            Left            =   5160
            TabIndex        =   16
            Top             =   375
            Width           =   240
         End
         Begin VB.CheckBox Check2 
            Height          =   240
            Index           =   7
            Left            =   4650
            TabIndex        =   15
            Top             =   375
            Width           =   240
         End
         Begin VB.CheckBox Check2 
            Height          =   240
            Index           =   6
            Left            =   4150
            TabIndex        =   14
            Top             =   375
            Width           =   240
         End
         Begin VB.CheckBox Check2 
            Height          =   240
            Index           =   5
            Left            =   3650
            TabIndex        =   13
            Top             =   375
            Width           =   240
         End
         Begin VB.CheckBox Check2 
            Height          =   240
            Index           =   4
            Left            =   3150
            TabIndex        =   12
            Top             =   375
            Width           =   240
         End
         Begin VB.CheckBox Check2 
            Height          =   240
            Index           =   3
            Left            =   2650
            TabIndex        =   11
            Top             =   375
            Width           =   240
         End
         Begin VB.CheckBox Check2 
            Height          =   240
            Index           =   2
            Left            =   2125
            TabIndex        =   10
            Top             =   375
            Width           =   240
         End
         Begin VB.CheckBox Check2 
            Height          =   240
            Index           =   1
            Left            =   1625
            TabIndex        =   9
            Top             =   375
            Width           =   240
         End
         Begin VB.CheckBox Check2 
            Height          =   240
            Index           =   0
            Left            =   1125
            TabIndex        =   8
            Top             =   375
            Width           =   240
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   2
            Left            =   1125
            MaxLength       =   4
            TabIndex        =   7
            Top             =   0
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BE"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   10
            Left            =   6165
            TabIndex        =   46
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BI"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   9
            Left            =   5660
            TabIndex        =   45
            Top             =   600
            Width           =   210
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C6"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   5
            Left            =   3645
            TabIndex        =   44
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C9"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   8
            Left            =   5160
            TabIndex        =   43
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C8"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   7
            Left            =   4650
            TabIndex        =   42
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C7"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   6
            Left            =   4155
            TabIndex        =   39
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C5"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   4
            Left            =   3150
            TabIndex        =   38
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C4"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   3
            Left            =   2650
            TabIndex        =   37
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C3"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   2
            Left            =   2100
            TabIndex        =   36
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C2"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   1605
            TabIndex        =   35
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   0
            Left            =   1125
            TabIndex        =   34
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad"
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
            Index           =   6
            Left            =   60
            TabIndex        =   33
            Top             =   75
            Width           =   765
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   3
         Left            =   7065
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3445
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   7425
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2965
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Demora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   2
         Top             =   1800
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C1DBD8&
         Caption         =   "Cancelar"
         Height          =   450
         Index           =   1
         Left            =   8175
         TabIndex        =   20
         Top             =   4360
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00CCC8AC&
         Caption         =   "Grabar"
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
         Height          =   570
         Index           =   0
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4210
         Width           =   1290
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3445
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2965
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   4
         Top             =   2965
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   1035
         Index           =   0
         Left            =   1200
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   375
         Width           =   8055
      End
      Begin VB.Label Label1 
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
         Height          =   195
         Index           =   8
         Left            =   650
         TabIndex        =   47
         Top             =   2460
         Width           =   330
      End
      Begin VB.Label Label1 
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
         Height          =   195
         Index           =   7
         Left            =   420
         TabIndex        =   41
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   7875
         TabIndex        =   40
         Top             =   1425
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Index           =   5
         Left            =   6045
         TabIndex        =   31
         Top             =   3445
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Index           =   4
         Left            =   6675
         TabIndex        =   30
         Top             =   3010
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   840
         TabIndex        =   29
         Top             =   4405
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lCodAlfa 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2490
         TabIndex        =   28
         Top             =   4405
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Index           =   3
         Left            =   375
         TabIndex        =   27
         Top             =   3505
         Width           =   570
      End
      Begin VB.Label Label1 
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
         Height          =   195
         Index           =   2
         Left            =   2040
         TabIndex        =   26
         Top             =   3010
         Width           =   660
      End
      Begin VB.Label Label1 
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
         Height          =   195
         Index           =   1
         Left            =   675
         TabIndex        =   25
         Top             =   3010
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Novedad"
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
         Left            =   300
         TabIndex        =   24
         Top             =   375
         Width           =   780
      End
   End
End
Attribute VB_Name = "RNov1c_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mMovil(3) As String
Dim mCodAlfa As String
Dim mKm As String
Dim mCodigoReferencia As Integer
Dim mDescr As String
Dim mCodNov As String
Dim mClima As String
Dim mImgNov As String
Dim mWhere As String
Dim mDescrInter As String
Dim mCodInter As Integer
Dim mResp As Boolean
Dim mPc As String

Private Sub Form_Load()
Dim mObj As New clRNov
Dim mI As Integer
Dim mRec As New ADODB.Recordset

   mPc = Mid(MDI.mPCname, 1, Len(MDI.mPCname) - 1)
   Me.Height = 5200
   Me.Width = 9555
'   Me.Height = 3015
'   Me.Width = 8220
   Me.Top = RNov1a_frm.Top + RNov1a_frm.Height + 30
   Me.Left = (MDI.Width - Me.Width) / 2

   Set mRec = mObj.oTabla("ramales", "")
   Do While Not mRec.EOF
     Combo1(4).AddItem mRec!descripcion & Space(50) & mRec!Abrevia & Space(2) & mRec!Codigo
     mRec.MoveNext
   Loop
   mRec.Close

   Set mRec = mObj.oTabla("origen", "WHERE Fecha_Baja IS NULL")
   Do While Not mRec.EOF
     Combo1(1).AddItem mRec!Codigo & " - " & mRec!descripcion
     mRec.MoveNext
   Loop
   mRec.Close
   For mI = 0 To 2
     mMovil(mI) = ""
   Next
   Check1.Visible = Not RNov1a_frm.List1.Visible
   For mI = 0 To 4
      RNov1a_frm.Command2(mI).Enabled = False
   Next
   RNov1d_frm.Enabled = False
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Private Sub Combo1_Click(Index As Integer)


Dim mI As Integer
   
   Select Case Index
      Case 1
         If Label1(3).Caption = "Cód. Alfa" Then
            'Text1(1).Text = Mid(Left(Trim(Combo1(Index).Text), InStr(1, Trim(Combo1(Index).Text), " ") - 1), 11)
            For mI = 0 To Combo1(4).ListCount - 1
               If Left(Right(Combo1(4).List(mI), 6), 2) = Left(Right(Trim(Combo1(Index).Text), 3), 2) Then
                  Combo1(4).ListIndex = mI
               End If
            Next
            For mI = 0 To Combo1(0).ListCount - 1
               'If Left(Combo1(0).List(mI), 2) = Mid(Trim(Combo1(Index).Text), 17, 2) Then 'mp 20160315
               If Left(Combo1(0).List(mI), 2) = Left(Right(Trim(Combo1(Index).Text), 7), 2) Then
                  Combo1(0).ListIndex = mI
               End If
            Next
            For mI = 0 To Combo1(5).ListCount - 1
               If Trim(Right(Combo1(5).List(mI), 3)) = fGetCodigoReferencia(Mid(Combo1(Index).Text, 2, 7)) Then
                  Combo1(5).ListIndex = mI
               End If
            Next
            
            Text1(1).Text = Mid(Left(Trim(Combo1(Index).Text), InStr(1, Trim(Combo1(Index).Text), " ") - 1), 11)
            Text1(1).Text = fGetKm(Mid(Combo1(Index).Text, 2, 7))
            
         End If
         Frame2.Visible = (InStr(1, Combo1(Index).Text, "RETIR", vbTextCompare) > 0)
  Case 4
         If Combo1(4).ListIndex >= 0 Then
            sLlenoSentido
            sLlenoReferencia
         End If
  Case 5
         If Combo1(5).ListIndex >= 0 Then
            Dim mCodReferencia As String
            mCodReferencia = Right(Combo1(5).Text, 3)
            sCompletaKM mCodReferencia
         End If
  End Select
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 27 Then
      Combo1(Index).ListIndex = -1
   End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mObj As New clRNov
Dim mI As Integer
Dim mJ As Integer
Dim xObjMov As Object
Dim xObjLbl As Object
Dim mCodNextel As String
Dim mError As Boolean
Dim mTipoNov As String
Dim mTexto As String
Dim mRec As New ADODB.Recordset

   mWhere = ""
   mTipoNov = ""
   
   If Combo1(5).ListIndex = -1 Then
      mCodigoReferencia = 0 'Referencia nula
   Else
      mCodigoReferencia = Trim(Right(Combo1(5).Text, 3))
   End If
   
   If Index = 0 Then
      mError = False
      Select Case Frame1.Caption
         Case "Ingreso de Novedades"
            If Text1(0).Text <> "" And Progresiva_Ok(Trim(Text1(1).Text), Trim(Right(Combo1(0).Text, 2))) And Combo1(0).Text <> "" And Combo1(1).Text <> "" Then
               If lCodAlfa.Visible Then  'Si la novedad proviene desde el móvil
                  mCodAlfa = lCodAlfa.Caption
                  If RNov1a_frm.List1.Visible Then
                     mMovil(0) = RNov1a_frm.List1.Text
                  End If
                  lCodAlfa.Visible = False
               Else
                  If RNov1a_frm.List1.Visible = True Or Check1.Value = 1 Then
                     mCodAlfa = fNewCodAlfa
                  Else
                     mCodAlfa = ""
                  End If
               End If
               mKm = Format(Trim(Text1(1).Text), "00.00")
               'mKm = Replace(Trim(Text1(1).Text), ".", ",") 'mp 20160309
               mKm = Trim(Text1(1).Text)
               mClima = ClimaOK(mKm)
               '*********************************************
               'Asignación de Evento
               If RNov1a_frm.List1.Visible = True Then
                  For mI = 0 To RNov1a_frm.List1.ListCount - 1
                     RNov1a_frm.List1.ListIndex = mI
                     mMovil(mI) = RNov1a_frm.List1.Text
                  Next
                  RNov1a_frm.Label3.Visible = False
                  RNov1a_frm.List1.Visible = False
                  For mI = 0 To RNov1a_frm.List1.ListCount - 1
                     mWhere = mWhere & " Codigo='" & mMovil(mI) & "' OR "
                     If Left(mMovil(mI), 1) = "M" Then  'Patrullas
                        sCambioEstados RNov1b_frm.Pat, RNov1b_frm.Label3, mCodAlfa, "p8", mKm, mI
                        mImgNov = "p8"
                        mCodNov = "A"
                     Else
                        mCodNov = "MM"
                        If Left(mMovil(mI), 2) = "GP" Then 'Grúas Pesadas
                           sCambioEstados RNov1b_frm.GPesada, RNov1b_frm.Label7, mCodAlfa, "x2", mKm, mI
                           mImgNov = "x2"
                        Else                               'Grúas
                           sCambioEstados RNov1b_frm.Grua, RNov1b_frm.Label4, mCodAlfa, "g6", mKm, mI
                           mImgNov = "g6"
                        End If
                     End If
                  Next
                  mImgNov = ""
                  mWhere = Mid(mWhere, 1, Len(mWhere) - 3)
                  mObj.xUpEstToolMov mWhere, "O", "(" & mCodAlfa & ")-" & mKm & " " & Left(Combo1(0).Text, 2) & " " & Left(Right(Combo1(4).Text, 7), 4)
                  mDescr = Trim(Text1(0).Text)
                  mCodNov = "A"
               Else ' Es una novedad desde un móvil
                  If Label2.Visible Then
                     mMovil(0) = Trim(Label2.Caption)
                     If Left(Label2.Caption, 1) = "M" Then
                        If RNov1b_frm.Label3(RNov1b_frm.mIndexMov).Tag = "p8" Then
                           sCambioEstados RNov1b_frm.Pat, RNov1b_frm.Label3, mCodAlfa, "p5", mKm, 0
                        End If
                     End If
                     If Left(Label2.Caption, 2) = "GP" Then
                        If RNov1b_frm.Label7(RNov1b_frm.mIndexMov).Tag = "x2" Then
                           sCambioEstados RNov1b_frm.GPesada, RNov1b_frm.Label7, mCodAlfa, "x3", mKm, 0
                        End If
                     End If
                     If Left(Label2.Caption, 2) = "G0" Then
                        If RNov1b_frm.Label4(RNov1b_frm.mIndexMov).Tag = "g6" Then
                           sCambioEstados RNov1b_frm.Grua, RNov1b_frm.Label4, mCodAlfa, "g2", mKm, 0
                        End If
                     End If
                     mDescr = Label2.Caption & " - " & Trim(Text1(0).Text)
                  Else
                     mDescr = Trim(Text1(0).Text)
                  End If
                  If Check1.Value = 1 Then
                     mCodNov = "D"
                  Else
                     mCodNov = "N"
                  End If
               End If
                mResp = mObj.xInsNovedades("", mCodAlfa, Trim(Right(MDI.mUser, 15)), mKm, Right(Combo1(0).Text, 2), Left(Combo1(1).Text, 3), mDescr, mCodNov, "0", mMovil(0), "", 0, mMovil(1), "", 0, mMovil(2), "", 0, "", "", Right(Combo1(4).Text, 1), mCodigoReferencia)
                mObj.xUpActualizarNot mPc, 1
               fInitRNov1a_frm
            Else
               mError = True
            End If
            
         Case "Asignar Tareas" ' Otros para Patrullas"
            If Progresiva_Ok(Text1(1).Text, Trim(Right(Combo1(0).Text, 2))) And Combo1(0).ListIndex <> -1 And Combo1(1).ListIndex <> -1 Then
               If lCodAlfa.Visible Then 'Para cuando lo asigno desde FlexGrid
                  mCodAlfa = lCodAlfa.Caption
                  mCodNov = "R"
               Else
                  mCodAlfa = fNewCodAlfa
                  mCodNov = "E"
               End If
               'mKm = Format(Text1(1).Text, "00.00")
               mKm = (Text1(1).Text) 'mp 20160309
               mClima = ClimaOK(mKm)
               '///// Copia en vector los moviles seleccionados en el LIST
               If RNov1a_frm.List1.Visible = True Then
                  For mI = 0 To RNov1a_frm.List1.ListCount - 1
                     RNov1a_frm.List1.ListIndex = mI
                     mMovil(mI) = RNov1a_frm.List1.Text
                  Next
                  RNov1a_frm.Label3.Visible = False
                  RNov1a_frm.List1.Visible = False
               End If
               mDescr = Combo1(1).Text & " - Móvil " & mMovil(0)
               '///// Fin de la copia
               If Frame2.Visible = True And (Trim(Text1(2).Text) = "" Or (Check2(0).Value = 0 And Check2(1).Value = 0 _
                  And Check2(2).Value = 0 And Check2(3).Value = 0 And Check2(4).Value = 0 And Check2(5).Value = 0 _
                  And Check2(6).Value = 0)) Then
                  MsgBox "Falta ingresar cantidad o carril.", vbCritical, sMessage
                  mError = True
               Else
                  mTexto = " "
                  If Check2(0).Value = 1 Then mTexto = mTexto & "C1 "
                  If Check2(1).Value = 1 Then mTexto = mTexto & "C2 "
                  If Check2(2).Value = 1 Then mTexto = mTexto & "C3 "
                  If Check2(3).Value = 1 Then mTexto = mTexto & "C4 "
                  If Check2(4).Value = 1 Then mTexto = mTexto & "C5 "
                  If Check2(5).Value = 1 Then mTexto = mTexto & "C6 "
                  If Check2(6).Value = 1 Then mTexto = mTexto & "C7 "
                  If Check2(7).Value = 1 Then mTexto = mTexto & "C8 "
                  If Check2(8).Value = 1 Then mTexto = mTexto & "C9 "
                  If Check2(9).Value = 1 Then mTexto = mTexto & "BI "
                  If Check2(10).Value = 1 Then mTexto = mTexto & "BE "
                  mTexto = mTexto & " - Cant: " & Trim(Text1(2).Text)
                  mResp = mObj.xInsNovedades("", mCodAlfa, Trim(Right(MDI.mUser, 15)), mKm, Right(Combo1(0).Text, 2), Left(Combo1(1).Text, 3), mDescr & mTexto, mCodNov, "0", mMovil(0), "", 0, "", "", 0, "", "", 0, "", "", Right(Combo1(4).Text, 1), mCodigoReferencia)
                  'CARGO CARRILES DE TAREA Y CANTIDAD DE RETIRO
                  If Check2(0).Visible = True Then
                     On Error Resume Next
                        mObj.xInTareaDatos mCodAlfa, mMovil(0), Check2(0).Value, Check2(1).Value, Check2(2).Value, Check2(3).Value, Check2(4).Value, Check2(5).Value, Check2(6).Value, Check2(7).Value, Check2(8).Value, Check2(9).Value, Check2(10).Value, Trim(Text1(2).Text)
                  End If
                  mObj.xUpActualizarNot mPc, 1
                  mObj.xUpEstMoviles mMovil(0), "O", "(" & mCodAlfa & ")-" & mKm & " " & Left(Combo1(0).Text, 2) & " " & Left(Right(Combo1(4).Text, 7), 4), "p2" 'mp 20160414
                  Unload RNov1b_frm
                  RNov1b_frm.Show
                  fInitRNov1a_frm
               End If
            Else
               'MsgBox "Faltan Datos", vbCritical, sMessage
               mError = True
            End If
       
         Case "Novedad de Tareas"
            If Text1(0).Text <> "" And Text1(1).Text <> "" And Combo1(0).ListIndex <> -1 Then
               mDescr = Trim(Text1(0).Text)
               'mKm = Format(Text1(1).Text, "00.00")
               mKm = Trim(Text1(1).Text)  'mp 20160309
               mClima = ClimaOK(mKm)
               mCodAlfa = lCodAlfa.Caption
               mMovil(0) = Label2.Caption
               mResp = mObj.xInsNovedades("", mCodAlfa, Trim(Right(MDI.mUser, 15)), mKm, Right(Combo1(0).Text, 2), Left(Combo1(1).Text, 3), mDescr, "N", "0", mMovil(0), "", 0, "", "", 0, "", "", 0, "", "", Right(Combo1(4).Text, 1), mCodigoReferencia)
               mObj.xUpActualizarNot mPc, 1
               mObj.xUpToolMov "codigo='" & mMovil(0) & "'", " tooltip=concat_ws('', LEFT(tooltip,10),'" & mKm & " " & Left(Combo1(0).Text, 2) & " " & Left(Right(Combo1(4).Text, 7), 4) & "')"
               fInitRNov1a_frm
               Unload RNov1b_frm
               RNov1b_frm.Show
            Else
               mError = True
            End If
            
         Case "Asignar Rutinas"
            If Progresiva_Ok(Text1(1).Text, Trim(Right(Combo1(0).Text, 2))) And Combo1(0).ListIndex <> -1 And Combo1(1).ListIndex <> -1 Then
               mCodAlfa = fNewCodAlfa
               '///// Copia en vector los moviles seleccionados en el LIST
               If RNov1a_frm.List1.Visible = True Then
                  RNov1a_frm.List1.ListIndex = 0
                  mMovil(0) = RNov1a_frm.List1.Text
                  RNov1a_frm.Label3.Visible = False
                  RNov1a_frm.List1.Visible = False
               End If
               mDescr = Combo1(1).Text & " - Móvil " & mMovil(0)
               'mKm = Format(Text1(1).Text, "00.00")
               mKm = Trim(Text1(1).Text)  'mp 20160309
               mClima = ClimaOK(mKm)
               '///// Fin de la copia
               mResp = mObj.xInsNovedades("", mCodAlfa, Trim(Right(MDI.mUser, 15)), mKm, Right(Combo1(0).Text, 2), Left(Combo1(1).Text, 3), mDescr, "MN", mClima, mMovil(0), "", 0, "", "", 0, "", "", 0, "", "", Right(Combo1(4).Text, 1), mCodigoReferencia)
               mObj.xUpActualizarNot mPc, 1
               'mObj.xUpEstMoviles mMovil(0), "O", "(" & mCodAlfa & ")-" & Format(Text1(1).Text, "00.00") & " " & Left(Combo1(0).Text, 2) & " " & Left(Right(Combo1(4).Text, 7), 4), "p7"
               mObj.xUpEstMoviles mMovil(0), "O", "(" & mCodAlfa & ")-" & mKm & " " & Left(Combo1(0).Text, 2) & " " & Left(Right(Combo1(4).Text, 7), 4), "p7"
               Unload RNov1b_frm
               RNov1b_frm.Show
               fInitRNov1a_frm
            Else
               MsgBox "Faltan Datos", vbCritical, "Versión Test"
               mError = True
            End If
               
         Case "Retome de Móvil"
            If Progresiva_Ok(Text1(1).Text, Trim(Right(Combo1(0).Text, 2))) And Combo1(0).ListIndex <> -1 Then
               'mKm = Format(Text1(1).Text, "00.00")
               mKm = Trim(Text1(1).Text) 'mp 20160309
               mClima = ClimaOK(mKm)
               mCodAlfa = lCodAlfa.Caption
               mDescr = "Retome Móvil " & Label2.Caption
               mResp = mObj.xInsNovedades("", mCodAlfa, Trim(Right(MDI.mUser, 15)), mKm, Right(Combo1(0).Text, 2), Left(Combo1(1).Text, 3), mDescr, "K", mClima, Label2.Caption, "", 0, "", "", 0, "", "", 0, "", "", Right(Combo1(4).Text, 1), mCodigoReferencia)
               mObj.xUpActualizarNot mPc, 1
            Else
               mError = True
            End If
             
         Case "Pedido de Móvil", "Pedido de Móvil por Demora"
            If Text1(0).Text <> "" And Progresiva_Ok(Trim(Text1(1).Text), Trim(Right(Combo1(0).Text, 2))) And Combo1(0).Text <> "" And Combo1(1).Text <> "" Then
               mMovil(0) = RNov1a_frm.List1.List(0)
               RNov1a_frm.Label3.Visible = False
               RNov1a_frm.List1.Visible = False
               mCodAlfa = lCodAlfa.Caption
               'mKm = Format(Trim(Text1(1).Text), "00.00")
               mKm = Trim(Text1(1).Text) 'mp 20160309
               mClima = ClimaOK(mKm)
               If Left(mMovil(0), 1) = "M" Then  'Patrullas
                  sCambioEstados RNov1b_frm.Pat, RNov1b_frm.Label3, mCodAlfa, "p8", mKm, 0
                  mObj.waze_asignar_movil mCodAlfa 'MP20171116
               Else
                  If Left(mMovil(0), 2) = "GP" Then 'Grúas Pesadas
                     sCambioEstados RNov1b_frm.GPesada, RNov1b_frm.Label7, mCodAlfa, "x2", mKm, 0
                  Else                               'Grúas
                     sCambioEstados RNov1b_frm.Grua, RNov1b_frm.Label4, mCodAlfa, "g6", mKm, 0
                  End If
               End If
               mObj.xUpEstToolMov "Codigo='" & mMovil(0) & "'", "O", "(" & mCodAlfa & ")-" & mKm & " " & Left(Combo1(0).Text, 2) & " " & Left(Right(Combo1(4).Text, 7), 4)
               mDescr = Trim(Text1(0).Text)
               mCodNov = "MM"
               mResp = mObj.xInsNovedades("", mCodAlfa, Trim(Right(MDI.mUser, 15)), mKm, Right(Combo1(0).Text, 2), Left(Combo1(1).Text, 3), mDescr, mCodNov, mClima, mMovil(0), "", 0, "", "", 0, "", "", 0, "", "", Right(Combo1(4).Text, 1), mCodigoReferencia)
               mObj.xUpActualizarNot mPc, 1
               If Frame1.Caption = "Pedido de Móvil por Demora" Then
                   mObj.xUpNovedadesSet "codnov='Q', Demora=NOW()", "fecha='" & Format(lCodAlfa.Tag, "yyyy-mm-dd hh:mm:ss") & "' AND codigo='" & lCodAlfa.Caption & "'"
                   lCodAlfa.Tag = ""
                   RNov1a_frm.Label4.Visible = mObj.bExistDatoTabla("novedades2", "codnov='D'")
               Else
                  If Len(lCodAlfa.Tag) > 5 Then
                     mObj.xUpNovedadesSet "codigo='" & lCodAlfa.Caption & "'", "fecha='" & Format(Left(lCodAlfa.Tag, 19), "yyyy-mm-dd hh:mm:ss") & "' AND Km='" & Right(lCodAlfa.Tag, 5) & "'"
                     lCodAlfa.Tag = ""
                  End If
               End If
               fInitRNov1a_frm
            End If
             
         Case "Pedido de Móviles"
            If Text1(0).Text <> "" And Progresiva_Ok(Trim(Text1(1).Text), Trim(Right(Combo1(0).Text, 2))) And Combo1(0).Text <> "" And Combo1(3).Text <> "" Then
               mMovil(0) = Label2.Caption
               If lCodAlfa.Visible Then
                  mCodAlfa = lCodAlfa.Caption
               Else
                  If Combo1(1).Text = "" Then
                     mCodAlfa = fNewCodAlfa
                  Else
                     mCodAlfa = Mid(Combo1(1).Text, 2, 7)
                  End If
               End If
               'mKm = Format(Trim(Text1(1).Text), "00.00")
               mKm = Trim(Text1(1).Text) 'mp 20160309
               mClima = ClimaOK(mKm)
               Select Case Label2.Caption
                  Case "AMBU"
                     mImgNov = "8"
                     mTipoNov = Left(Combo1(3).Text, 1)
                  Case "GEND"
                     mImgNov = "2"
                  Case "BOMB"
                     mImgNov = "4"
                     mTipoNov = Left(Combo1(3).Text, 1)
                  Case "POLI"
                     mImgNov = "6"
                  Case "AMB1" 'AMBU GCO
                     mImgNov = "8"
                     mTipoNov = Left(Combo1(3).Text, 1)
               End Select
               sCambioEstados RNov1b_frm.MovExternos, RNov1b_frm.Label8, mCodAlfa, mImgNov, mKm, 0
               Set mRec = mObj.oTabla("moviles", "WHERE Codigo='" & mMovil(0) & "'")
               If Not mRec.EOF Then
                  If mRec!ToolTip = "" Then
                     mObj.xUpEstMoviles mMovil(0), "O", "(" & mCodAlfa & ")-" & mKm & " " & Left(Combo1(0).Text, 2) & " " & Left(Right(Combo1(4).Text, 7), 4), mImgNov
                  Else
                     mObj.xUpToolMov "Codigo='" & mMovil(0) & "'", "ToolTip = concat_ws('&',ToolTip,'(" & mCodAlfa & ")-" & mKm & " " & Left(Combo1(0).Text, 2) & " " & Left(Right(Combo1(4).Text, 7), 4) & "'), Estado='O', CodNov='" & mImgNov & "'"
                  End If
               End If
               mDescr = Trim(Text1(0).Text)
               mCodNov = "M"
               mTipoNov = Left(Combo1(3).Text, 1)
               mResp = mObj.xInsNovedades("", mCodAlfa, Trim(Right(MDI.mUser, 15)), mKm, Right(Combo1(0).Text, 2), "SSV", mDescr, mCodNov, mClima, mMovil(0), Left(Combo1(2).Text, 1), 0, "", "", 0, "", "", 0, "", mTipoNov, Right(Combo1(4).Text, 1), mCodigoReferencia)
               mObj.xUpActualizarNot mPc, 1
               Unload RNov1b_frm
               RNov1b_frm.Show
            Else
               mError = True
            End If
             
         Case "Arribo de Móviles"
            If Text1(0).Text <> "" And Progresiva_Ok(Text1(1).Text, Trim(Right(Combo1(0).Text, 2))) And Combo1(0).ListIndex > -1 And Combo1(1).ListIndex > -1 Then
               sProcMovExt
               mDescr = Label2.Caption & " - Arribo Móvil " & Trim(Text1(0).Text)
               mResp = mObj.xInsNovedades("", mCodAlfa, Trim(Right(MDI.mUser, 15)), mKm, Right(Combo1(0).Text, 2), "SSV", mDescr, "L", mClima, Label2.Caption, "", 0, "", "", 0, "", "", 0, "", "", Right(Combo1(4).Text, 1), mCodigoReferencia)
               mObj.xUpActualizarNot mPc, 1
               Unload RNov1b_frm
               RNov1b_frm.Show
            Else
              mError = True
            End If
              
         Case "Cancelar Pedido de Móvil"
            If Progresiva_Ok(Text1(1).Text, Trim(Right(Combo1(0).Text, 2))) And Combo1(0).ListIndex > -1 And Combo1(1).ListIndex > -1 Then
               sProcMovExt
               mDescr = Label2.Caption & " - Pedido de Móvil Cancelado"
               mResp = mObj.xInsNovedades("", mCodAlfa, Trim(Right(MDI.mUser, 15)), mKm, Right(Combo1(0).Text, 2), Left(Combo1(1).Text, 3), mDescr, "T", mClima, Label2.Caption, "", 0, "", "", 0, "", "", 0, "", mTipoNov, Right(Combo1(4).Text, 1), mCodigoReferencia)
               mObj.xUpActualizarNot mPc, 1
               Unload RNov1b_frm
               RNov1b_frm.Show
            Else
               mError = True
            End If
            
         Case "Móvil No Arribó"
            If Progresiva_Ok(Text1(1).Text, Trim(Right(Combo1(0).Text, 2))) And Combo1(0).ListIndex > -1 And Combo1(1).ListIndex > -1 Then
               sProcMovExt
               mDescr = Label2.Caption & " - Móvil NO ARRIBO"
               mResp = mObj.xInsNovedades("", mCodAlfa, Trim(Right(MDI.mUser, 15)), mKm, Right(Combo1(0).Text, 2), Left(Combo1(1).Text, 3), mDescr, "V", mClima, Label2.Caption, "", 0, "", "", 0, "", "", 0, "", mTipoNov, Right(Combo1(4).Text, 1), mCodigoReferencia)
               mObj.xUpActualizarNot mPc, 1
               Unload RNov1b_frm
               RNov1b_frm.Show
            Else
               mError = True
            End If
         
      End Select
   Else 'CANCELO EL FORMULARIO
      fInitRNov1a_frm
      RNov1a_frm.Label3.Visible = False
   End If
   If Not mError Then
      If RNov1a_frm.Image1.Visible Then
        RNov1a_frm.Image1.Visible = False
        mObj.xUpActualizarNot "-", 0
        Unload RNov1b_frm
        RNov1b_frm.Show
      End If
      Unload RNov1d_frm
      RNov1d_frm.Show
      Set mObj = Nothing
      Unload RNov1c_frm
   Else
      MsgBox "Completar todos los campos.", vbCritical, sMessage
      Set mObj = Nothing
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim mI As Integer
   RNov1a_frm.Enabled = True
   For mI = 0 To 4
      RNov1a_frm.Command2(mI).Enabled = True
   Next
   RNov1b_frm.Enabled = True
   RNov1d_frm.Enabled = True
End Sub


Private Sub Text1_Change(Index As Integer)
   If Index = 0 Then
      Label4.Caption = Len(Text1(0).Text) & " / 255"
   End If
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 1 Then 'Progresiva
      If KeyAscii <> 46 Then
         KeyAscii = fNumeroKeyPress(KeyAscii)
      End If
   Else 'Novedad
      KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
      KeyAscii = fUcaseKeyPress(KeyAscii)
   End If
End Sub

Public Sub sInitTareas()
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
   Frame1.Caption = "Asignar Tareas"
   Label1(0).Visible = False
   Check1.Visible = False
   Text1(0).Visible = False
   Label1(3).Caption = "Tareas"
   Combo1(1).Clear
   Set mRec = mObj.oTabla("otros", "")
   Do While Not mRec.EOF
     Combo1(1).AddItem mRec!Codigo & " - " & mRec!descripcion
     mRec.MoveNext
   Loop
   mRec.Close
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Public Sub sInitRutinas()
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
  Frame1.Caption = "Asignar Rutinas"
  Label1(0).Visible = False
  Check1.Visible = False
  Text1(0).Visible = False
  Label1(3).Caption = "Rutinas"
  Combo1(1).Clear
  Set mRec = mObj.oTabla("rutinas", "")
  Do While Not mRec.EOF
    Combo1(1).AddItem mRec!Codigo & " - " & mRec!descripcion
    mRec.MoveNext
  Loop
  mRec.Close
  Set mObj = Nothing
  Set mRec = Nothing
End Sub

Public Sub sInitMovExternos()
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset

   Check1.Visible = False
   Label1(3).Caption = "Cód. Alfa"
   Label1(5).Visible = True
   Combo1(1).Width = 2700
   Combo1(3).Visible = True
   Combo1(1).Clear
   Set mRec = mObj.oMovilesDistTool("O", "AND (CodTipoMov = 'PAT' OR CodTipoMov = 'GRU')")
   Do While Not mRec.EOF
      If Len(mRec!ToolTip) > 7 Then
         Combo1(1).AddItem mRec!ToolTip
      End If
      mRec.MoveNext
   Loop
   mRec.Close
   If Label2.Caption = "AMBU" Or Label2.Caption = "AMB1" Then
      Label1(4).Visible = True
      Combo1(2).Visible = True
      Combo1(2).Clear
      Combo1(2).AddItem "1- AMARILLO"
      Combo1(2).AddItem "2- ROJO"
      Combo1(2).AddItem "3- VERDE"
      Combo1(2).ListIndex = 1
   End If
   Combo1(1).ListIndex = -1
   Combo1(3).AddItem "O-OTROS"
   Combo1(3).AddItem "P-PEAJES"
   Combo1(3).AddItem "A-ADMINISTRACIÓN"
   Combo1(3).ListIndex = 0
   Label2.Visible = True
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Public Sub sInitArriboMovExt(pMovil As String)
Dim mObj As New clRNov
Dim mLong As Integer
Dim mRec As New ADODB.Recordset
 
   Check1.Visible = False
   Label1(3).Caption = "Cód. Alfa"
   Label1(4).Visible = False
   Label1(5).Visible = False
   Combo1(1).Width = 2700
   Combo1(1).Clear
   Set mRec = mObj.oTabla("moviles", "WHERE ESTADO = 'O' AND Codigo='" & pMovil & "'")
   If Not mRec.EOF Then
'      For mLong = 1 To Len(mRec!ToolTip) Step 22
'         Combo1(1).AddItem Mid(mRec!ToolTip, mLong, 22)
'         mLong = mLong + 1
'      Next

   Dim cadena As String
   cadena = mRec!ToolTip
   Dim retInst As Integer
   Dim longCadena As Integer
   Do While InStr(1, cadena, "&") <> 0
      retInst = InStr(1, cadena, "&")
      Combo1(1).AddItem Mid(cadena, 1, retInst - 1)
      longCadena = Len(cadena)
      cadena = Mid(cadena, retInst + 1, longCadena - retInst)
   Loop
      Combo1(1).AddItem cadena
   
   
   

   End If
   
   
   
   mRec.Close
   Combo1(2).Visible = False
   Combo1(3).Visible = False
   Combo1(1).ListIndex = -1
   Label2.Visible = True
   Label2.Caption = pMovil
   Text1(0).Width = 1000
   Label1(0).Caption = "Nro. Móvil"
   Frame1.Caption = "Arribo de Móviles"
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Public Sub sNovTareas()
   Frame1.Caption = "Novedad de Tareas"
   Label1(3).Visible = False
   Check1.Visible = False
   Combo1(1).Visible = False
End Sub

Public Sub sRetomeMov()
   Frame1.Caption = "Retome de Móvil"
   Label1(0).Visible = False
   Text1(0).Visible = False
   Label1(3).Visible = False
   Check1.Visible = False
   Combo1(1).Visible = False
   Label2.Visible = True
End Sub

Private Sub sCambioEstados(pObjMov As Object, pObjLbl As Object, pCodAlfa As String, pImg As String, pKm As String, pInd As Integer)
Dim mObj As New clRNov
Dim mJ As Integer
   For mJ = 0 To pObjMov.UBound
      If pObjMov(mJ).Tag = mMovil(pInd) Then
         pObjMov(mJ).Picture = LoadPicture(App.Path & "\RegNovedades\Image\Iconos\" & pImg & ".gif")
         pObjMov(mJ).ToolTipText = "(" & pCodAlfa & ")-" & pKm & " " & Left(Combo1(0).Text, 2) & " " & Left(Right(Combo1(4).Text, 7), 4)
         pObjLbl(mJ).Tag = pImg
      End If
   Next
   mObj.xUpMovilesCodNov mMovil(pInd), pImg
   Set pObjMov = Nothing
   Set pObjLbl = Nothing
   Set mObj = Nothing
End Sub

Private Sub sProcMovExt()
Dim mObj As New clRNov
Dim mI As Integer
Dim mJ As Integer
 
   mCodAlfa = Mid(Combo1(1).Text, 2, 7)
   mKm = Trim(Text1(1).Text)
   mClima = ClimaOK(mKm)
   mWhere = ""
   For mI = 0 To Combo1(1).ListCount - 1
      If Mid(Combo1(1).List(mI), 2, 7) = mCodAlfa Then
         For mJ = mI + 1 To Combo1(1).ListCount - 1
            mWhere = mWhere & Combo1(1).List(mJ) & "&"
         Next
         mI = mJ
      Else
         mWhere = mWhere & Combo1(1).List(mJ) & "&"
      End If
   Next
   If mWhere <> "" Then
      mWhere = Mid(mWhere, 1, Len(mWhere) - 1)
      mObj.xUpToolMov "Codigo='" & Label2.Caption & "'", "ToolTip='" & mWhere & "'"
   Else
      Select Case Label2.Caption
        Case "AMBU"
             mImgNov = "7"
        Case "GEND"
             mImgNov = "1"
        Case "BOMB"
             mImgNov = "3"
        Case "POLI"
             mImgNov = "5"
        Case "AMB1" 'AMBU GCO 'verrrr
            If RNov1c_frm.Frame1.Caption = "Cancelar Pedido de Móvil" Then
               mImgNov = "a1"
            Else
               mImgNov = "a2"  'ARRIBO
            End If
     End Select
      If Label2.Caption <> "AMB1" Then
         mObj.xUpEstMoviles Label2.Caption, "L", mWhere, mImgNov
      Else
         If RNov1c_frm.Frame1.Caption = "Cancelar Pedido de Móvil" Then
            mObj.xUpEstMoviles Label2.Caption, "L", mWhere, mImgNov
         Else
            mObj.xUpEstMoviles Label2.Caption, "O", RNov1b_frm.MovExternos(4).ToolTipText, mImgNov
         End If
      End If
   End If
   Set mObj = Nothing
End Sub
Private Sub sLlenoSentido()
Dim mCodRamal As String
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
   
   mCodRamal = Right(Combo1(4).Text, 1)
   Combo1(0).Clear
   Set mRec = mObj.oTabla("sentidos", "where codRamal = " & mCodRamal & " order by 2")
   Do While Not mRec.EOF
     Combo1(0).AddItem mRec!descripcion & Space(60) & mRec!Codigo
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
   
   mCodRamal = Right(Combo1(4).Text, 1)
   Combo1(5).Clear
   Set mRec = mObj.oTabla("referencias", "where codRamal = " & mCodRamal & " order by 2")
   Do While Not mRec.EOF
     Combo1(5).AddItem mRec!descripcion & Space(100) & mRec!Codigo
     mRec.MoveNext
   Loop
   mRec.Close
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Private Sub sCompletaKM(pCodReferencia As String)
   Dim mObj As New clRNov
   Text1(1).Text = mObj.sTablaDescr("referencias", "codigo=" & pCodReferencia, 2)
   Set mObj = Nothing
End Sub
