VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Peek2_frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8595
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   6240
      TabIndex        =   16
      Top             =   300
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid Flex1 
      Height          =   3135
      Left            =   60
      TabIndex        =   1
      Top             =   2520
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   5530
      _Version        =   327680
      Cols            =   5
      FixedCols       =   0
   End
   Begin VB.Label Label4 
      Height          =   255
      Index           =   3
      Left            =   1620
      TabIndex        =   15
      Top             =   2160
      Width           =   6855
   End
   Begin VB.Label Label4 
      Height          =   255
      Index           =   2
      Left            =   1620
      TabIndex        =   14
      Top             =   1800
      Width           =   6855
   End
   Begin VB.Label Label4 
      Height          =   255
      Index           =   1
      Left            =   1620
      TabIndex        =   13
      Top             =   1440
      Width           =   6855
   End
   Begin VB.Label Label4 
      Height          =   255
      Index           =   0
      Left            =   1620
      TabIndex        =   12
      Top             =   1080
      Width           =   6855
   End
   Begin VB.Label Label3 
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   11
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label3 
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   10
      Top             =   2160
      Width           =   500
   End
   Begin VB.Label Label3 
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   9
      Top             =   1800
      Width           =   500
   End
   Begin VB.Label Label3 
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   8
      Top             =   1440
      Width           =   500
   End
   Begin VB.Label Label3 
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   7
      Top             =   1080
      Width           =   500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   300
      TabIndex        =   6
      Top             =   720
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Km 47"
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
      Index           =   3
      Left            =   300
      TabIndex        =   5
      Top             =   2160
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Km 32"
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
      Index           =   2
      Left            =   300
      TabIndex        =   4
      Top             =   1800
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Km 23"
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
      Left            =   300
      TabIndex        =   3
      Top             =   1440
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Km 14"
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
      Left            =   300
      TabIndex        =   2
      Top             =   1080
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Estado de conexiones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2385
   End
End
Attribute VB_Name = "Peek2_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObj As New clPeek
Dim mRec As New ADODB.Recordset
Dim mI As Integer

Private Sub Form_Load()
sInitForm
End Sub

Private Sub sInitForm()
Dim mFecha As String
Dim mInsertFlex As String
With Flex1
   .Row = 0
   For mI = 0 To .Cols - 1
      .Col = mI
      .CellFontBold = True
   Next
   .ColWidth(0) = 2000   'Fecha
   .ColWidth(1) = 1200  'Km14
   .ColWidth(2) = 1200  'Km23
   .ColWidth(3) = 1200  'Km32
   .ColWidth(4) = 1200  'Km45
   .TextMatrix(0, 0) = "Fecha"
   .TextMatrix(0, 1) = "Km 14"
   .TextMatrix(0, 2) = "Km 23"
   .TextMatrix(0, 3) = "Km 32"
   .TextMatrix(0, 4) = "Km 45"
End With
Flex1.Tag = "0"
Set mRec = mObj.oFallas
If Not mRec.EOF Then
   mI = 1
   mFecha = mRec.Fields(1)
   Label3(4).Caption = mFecha
   Flex1.AddItem mFecha & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "", mI
   Flex1.RemoveItem 2
   Do While Not mRec.EOF
      If mFecha = mRec.Fields(1) Then
         Select Case mRec.Fields(0)
            Case "Km14"
               If Flex1.Tag = "0" Then
                  sTextos 0, mRec.Fields(2), NVL(mRec.Fields(3), "")
               End If
               Flex1.TextMatrix(mI, 1) = mRec.Fields(2)
            Case "Km23"
               If Flex1.Tag = "0" Then
                  sTextos 1, mRec.Fields(2), NVL(mRec.Fields(3), "")
               End If
               Flex1.TextMatrix(mI, 2) = mRec.Fields(2)
            Case "Km32"
               If Flex1.Tag = "0" Then
                  sTextos 2, mRec.Fields(2), NVL(mRec.Fields(3), "")
               End If
               Flex1.TextMatrix(mI, 3) = mRec.Fields(2)
            Case "Km47"
               If Flex1.Tag = "0" Then
                  sTextos 3, mRec.Fields(2), NVL(mRec.Fields(3), "")
               End If
               Flex1.TextMatrix(mI, 4) = mRec.Fields(2)
         End Select
         mRec.MoveNext
      Else
         Flex1.Tag = "1"
         mI = mI + 1
         Flex1.AddItem mFecha & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "", mI
         mFecha = mRec.Fields(1)
      End If
   Loop
End If
mRec.Close
End Sub

Private Sub sTextos(ByVal pInd As Integer, ByVal pEstado As String, ByVal pDescr As String)
Label3(pInd).Caption = pEstado
If pEstado = "KO" Then
   Label3(pInd).ForeColor = &HC0
Else
   Label3(pInd).ForeColor = &H0
End If
Label4(pInd).Caption = pDescr
End Sub
