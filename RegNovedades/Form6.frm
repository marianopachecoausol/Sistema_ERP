VERSION 5.00
Begin VB.Form RNov6_frm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificación de Novedad"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.ComboBox Combo1 
         BackColor       =   &H80000000&
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
         Index           =   1
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1560
         Width           =   3900
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H80000000&
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7780
         MaxLength       =   5
         TabIndex        =   13
         Top             =   1125
         Width           =   615
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   2850
         TabIndex        =   5
         Top             =   3480
         Width           =   3255
         Begin VB.CommandButton Command1 
            Caption         =   "&Cancelar"
            Height          =   495
            Index           =   1
            Left            =   1800
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Grabar"
            Default         =   -1  'True
            Height          =   495
            Index           =   0
            Left            =   480
            TabIndex        =   6
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox Text1 
         Height          =   1035
         Left            =   720
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   2280
         Width           =   7755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ramal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   17
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   7200
         TabIndex        =   15
         Top             =   2925
         Width           =   1290
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Left            =   4500
         TabIndex        =   12
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Left            =   1440
         TabIndex        =   11
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sentido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   3780
         TabIndex        =   10
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Km:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   7300
         TabIndex        =   9
         Top             =   1200
         Width           =   330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   3780
         TabIndex        =   8
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Modificación de Novedad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   2880
         TabIndex        =   4
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Left            =   750
         TabIndex        =   1
         Top             =   2040
         Width           =   780
      End
   End
End
Attribute VB_Name = "RNov6_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Click(Index As Integer)
   Select Case Index
     Case 0
         If Combo1(0).ListIndex >= 0 Then
            sLlenoSentido
         End If
   End Select
End Sub

Private Sub Form_Load()
Dim mObj As New clRNov
Dim mI As Integer
Dim mRec As New ADODB.Recordset

   sAlinearForm Me
   Set mRec = mObj.oTabla("ramales", "")
   Do While Not mRec.EOF
     Combo1(0).AddItem mRec!Descripcion & Space(50) & mRec!Abrevia & Space(2) & mRec!Codigo
     mRec.MoveNext
   Loop
   mRec.Close

   RNov1a_frm.Enabled = False
   RNov1b_frm.Enabled = False
   RNov1d_frm.Enabled = False

'Dim mI As Integer
'   sAlinearForm Me
'   Combo1(0).AddItem "A-Asc"
'   Combo1(0).AddItem "D-Des"
'   Combo1(0).AddItem "K-Col.Asc"
'   Combo1(0).AddItem "B-Col.Des"
'   Combo1(0).AddItem "S-Tro.Asc"
'   Combo1(0).AddItem "T-Tro.Des"
'   RNov1a_frm.Enabled = False
'   RNov1b_frm.Enabled = False
'   RNov1d_frm.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   RNov1a_frm.Enabled = True
   RNov1b_frm.Enabled = True
   RNov1d_frm.Enabled = True
   Unload RNov1d_frm
   RNov1d_frm.Show
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mObj As New clRNov
Dim mKm As String
Dim mSent As String
Dim mRamal As String
Dim mCod As String
Dim mFecha As String

   If Index = 0 Then
      mCod = Label4(0).Caption
      mFecha = Label4(1).Caption
      mKm = Text2.Text
      'mSent = Right(Combo1(1).Text, 2)
      'mRamal = Right(Combo1(0).Text, 1)
      
      'If Text1.Text <> "" And Combo1(1).Text <> "" And Text2.Text <> "" Then
      If Text1.Text <> "" And Text2.Text <> "" And Combo1(0).Text <> "" And Combo1(1).Text <> "" Then
         If Progresiva_Ok(Trim(Text2.Text), Trim(Right(Combo1(1).Text, 2))) Then
            'mObj.xUpNovedadesSet "Descripcion = '" & Trim(Text1.Text) & "', Km = " & Trim(Text2.Text) & ", Sent = " & Right(Combo1(1).Text, 2) & ", codramal = " & Right(Combo1(0).Text, 1), "Codigo = '" & mCod & "' AND Fecha = '" & Format(mFecha, "yyyy/mm/dd hh:mm:ss") & "' "
            mObj.xUpNovedadesSet "Descripcion = '" & Trim(Text1.Text) & "', Km = " & Replace(Trim(Text2.Text), ",", ".") & ", Sent = " & Right(Combo1(1).Text, 2) & ", codramal = " & Right(Combo1(0).Text, 1), "Codigo = '" & mCod & "' AND Fecha = '" & Format(mFecha, "yyyy/mm/dd hh:mm:ss") & "' " 'mp 20160309
            
            Unload RNov6_frm
            mObj.xUpActualizarNot Mid(Trim(MDI.PCname), 1, Len(Trim(MDI.PCname)) - 1), "1"
         Else
         End If
      Else
         'MsgBox "Debe Completar la Novedad", vbCritical, "RegNov 3.1  - Atención!"
         MsgBox "Completar todos los campos.", vbCritical, "RegNov 3.1  - Atención!"
      End If
      
   Else
      Unload RNov6_frm
   End If
   Set mObj = Nothing
End Sub

Private Sub Text1_Change()
   Label5.Caption = Len(Text1.Text) & " / 255"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
   KeyAscii = fUcaseKeyPress(KeyAscii)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 46 Then
      KeyAscii = fNumeroKeyPress(KeyAscii)
   End If
End Sub

Public Sub sLlenoSentido()
Dim mCodRamal As String
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
   
   mCodRamal = Right(Combo1(0).Text, 1)
   Combo1(1).Clear
   Set mRec = mObj.oTabla("sentidos", "where codRamal = " & mCodRamal & " order by 2")
   Do While Not mRec.EOF
     Combo1(1).AddItem mRec!Descripcion & Space(60) & mRec!Codigo
     mRec.MoveNext
   Loop
   mRec.Close
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

