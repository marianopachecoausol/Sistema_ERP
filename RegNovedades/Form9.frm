VERSION 5.00
Begin VB.Form RNov9_frm 
   Caption         =   "Módulo de Cambio de Atributo de una Novedad"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   10080
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   3855
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   9495
         Begin VB.Frame Frame4 
            Caption         =   "Atributos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   3480
            TabIndex        =   12
            Top             =   1680
            Width           =   2415
            Begin VB.OptionButton Option1 
               Caption         =   "Ninguno"
               Height          =   255
               Index           =   2
               Left            =   480
               TabIndex        =   15
               Top             =   1320
               Width           =   975
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Incidente"
               Height          =   255
               Index           =   1
               Left            =   480
               TabIndex        =   14
               Top             =   840
               Width           =   1215
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Accidente"
               Height          =   255
               Index           =   0
               Left            =   480
               TabIndex        =   13
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   735
            Left            =   1080
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   600
            Width           =   7095
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Volver"
         Height          =   495
         Index           =   1
         Left            =   5160
         TabIndex        =   5
         Top             =   4920
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Default         =   -1  'True
         Height          =   495
         Index           =   0
         Left            =   3360
         TabIndex        =   4
         Top             =   4920
         Width           =   1215
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
         Height          =   2790
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1920
         Width           =   9495
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   5775
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cambio de Atributos de Novedades"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   330
            Left            =   600
            TabIndex        =   7
            Top             =   165
            Width           =   4530
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
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
         Left            =   5280
         TabIndex        =   9
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
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
         Left            =   3360
         TabIndex        =   8
         Top             =   1080
         Width           =   1065
      End
   End
End
Attribute VB_Name = "RNov9_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mRec As New ADODB.Recordset

Private Sub Form_Load()
Me.Height = 6100
Me.Width = 10200
sAlinearForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mRec = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mObj As New clRNov
Dim xTypeNov As String
Dim xFecha As Date
Dim xNovedad As String
Dim mTipoNovAct As String
Dim mRamal As String
Dim mSent As String

If Index = 0 Then
   If Command1(0).Caption <> "Grabar" Then
      If sValidFechaDesdeHasta(Text1(0).Text, Text1(1).Text) Then
         List1.Clear
         Set mRec = mObj.oNovedadesFecha(Text1(0).Text, Text1(1).Text, "and codnov in ('A','B','O','M')", "", "ORDER BY FECHA")
         Do While Not mRec.EOF
            mRamal = mObj.sTablaDescr("ramales", "codigo =  " & mRec!codramal, 2)
            mSent = Left(mObj.sTablaDescr("sentidos", "codigo = " & mRec!sent, 1), 2)
            List1.AddItem mRec!Fecha & " - " & Left(mRec!km & Space(5), 5) & " - " & mSent & " - " & mRamal & " - " & Left(mRec!Descripcion & Space(80), 80) & "-" & Left(mRec!TipoNov & Space(2), 2)
            mRec.MoveNext
         Loop
         mRec.Close
      End If
   Else
      If Option1(0).Value Then 'Accidente
         xTypeNov = "X"
      End If
      If Option1(1).Value Then 'Incidente
         xTypeNov = "I"
      End If
      If Option1(2).Value Then
         xTypeNov = ""
      End If
      xFecha = Left(List1.Text, 19)
      mTipoNovAct = Trim(Right(List1.Text, 2))
      If Len(mTipoNovAct) = 2 Then
         xTypeNov = xTypeNov & Right(mTipoNovAct, 1) 'Cambio solo el primer caracter que indica si es Incidente o Accidente
      Else
         If mTipoNovAct = "O" Or mTipoNovAct = "P" Or mTipoNovAct = "A" Then
            xTypeNov = xTypeNov & mTipoNovAct
         End If
      End If
      mObj.xUpNovedadesSet "TipoNov = '" & xTypeNov & "'", "Fecha = '" & Format(xFecha, "yyyy/mm/dd hh:mm:ss") & "'"
      xNovedad = Left(List1.Text, 114)
      List1.RemoveItem List1.ListIndex
      List1.AddItem Left(xNovedad & Space(114), 114) & "-" & xTypeNov
      List1.Refresh
      Command1(0).Caption = "&Aceptar"
      Frame3.Visible = False
   End If
Else
   If Command1(0).Caption = "Grabar" Then
      Frame3.Visible = False
      Command1(0).Caption = "&Aceptar"
   Else
      Unload RNov9_frm
      ShowMenu 1, True, False
   End If
End If
Set mObj = Nothing
End Sub

Private Sub List1_DblClick()
Dim TipoNov As String
TipoNov = Trim(Right(List1.Text, 2))
If Len(TipoNov) = 2 Then
   TipoNov = Left(TipoNov, 1)
End If
If TipoNov = "X" Then
   Option1(0).Value = True
Else
   If TipoNov = "I" Then
      Option1(1).Value = True
   Else
      Option1(2).Value = True
   End If
End If
Text2.Text = Mid(List1.Text, 1, (Len(List1.Text) - 3))
Command1(0).Caption = "Grabar"
Frame3.Visible = True
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
End Sub


