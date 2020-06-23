VERSION 5.00
Begin VB.Form Peek_3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de estados de conexiones inalámbricas"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5640
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   6
      Left            =   1275
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Tag             =   "Km29"
      Top             =   1800
      Width           =   800
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E8E8E3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   6
      Left            =   2175
      MaxLength       =   100
      TabIndex        =   28
      Top             =   1800
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E8E8E3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   5
      Left            =   2160
      MaxLength       =   100
      TabIndex        =   25
      Top             =   3255
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E8E8E3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   4
      Left            =   2160
      MaxLength       =   100
      TabIndex        =   24
      Top             =   2895
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E8E8E3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   3
      Left            =   2160
      MaxLength       =   100
      TabIndex        =   23
      Top             =   2535
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E8E8E3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   2
      Left            =   2160
      MaxLength       =   100
      TabIndex        =   22
      Top             =   2175
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E8E8E3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
      Left            =   2160
      MaxLength       =   100
      TabIndex        =   21
      Top             =   1440
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E8E8E3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   0
      Left            =   2160
      MaxLength       =   100
      TabIndex        =   20
      Top             =   1080
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   5
      Left            =   1260
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Tag             =   "Rta5"
      Top             =   3255
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   4
      Left            =   1260
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Tag             =   "Km47"
      Top             =   2895
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   3
      Left            =   1260
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Tag             =   "Km36"
      Top             =   2535
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   2
      Left            =   1260
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Tag             =   "Km32"
      Top             =   2175
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   1
      Left            =   1260
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Tag             =   "Km23"
      Top             =   1440
      Width           =   800
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C1DBD8&
      Caption         =   "Salir"
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
      Index           =   1
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3915
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C1DBD8&
      Caption         =   "Grabar"
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
      Index           =   0
      Left            =   3180
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "A"
      Top             =   3915
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E8E8E3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   260
      Index           =   0
      Left            =   780
      MaxLength       =   10
      TabIndex        =   5
      Top             =   300
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   0
      Left            =   1260
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "Km14"
      Top             =   1080
      Width           =   800
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E8E8E3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   260
      Index           =   1
      Left            =   3000
      MaxLength       =   5
      TabIndex        =   0
      Top             =   300
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Km 29 -"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   6
      Left            =   375
      TabIndex        =   30
      Top             =   1860
      Width           =   690
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4560
      MouseIcon       =   "Peek_3.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   300
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
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
      Left            =   2160
      TabIndex        =   26
      Top             =   780
      Width           =   1140
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruta 5 -"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   5
      Left            =   360
      TabIndex        =   19
      Top             =   3315
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Km 45 -"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   18
      Top             =   2955
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Km 36 -"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   17
      Top             =   2595
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Km 32 -"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   16
      Top             =   2235
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Km 23 -"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   15
      Top             =   1500
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Km 14 -"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   14
      Top             =   1140
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   6
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
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
      Left            =   1260
      TabIndex        =   4
      Top             =   780
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hora"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contador"
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
      Left            =   180
      TabIndex        =   1
      Top             =   780
      Width           =   885
   End
End
Attribute VB_Name = "Peek_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mI As Integer

Private Sub Form_Load()
   sInitForm
   sAlinearForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ShowMenu 13, True, False
End Sub
   
Private Sub Combo1_Click(Index As Integer)
   Combo1(Index).ForeColor = vbBlue
   If Combo1(Index).Text = "KO" Then
      Combo1(Index).ForeColor = vbRed
   End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mObj As New clPeek
Dim mFlag As Boolean

   If Index = 0 Then
      If fValidar Then
         For mI = 0 To Combo1.UBound
            If Combo1(mI).Text <> "" Then
               mFlag = mObj.xInsFallasRed(Combo1(mI).Tag, Text1(0).Text & " " & Text1(1).Text, Combo1(mI).Text, Trim(Text2(mI).Text))
            End If
            
         Next
         MsgBox "Alta exitosa", vbInformation, sMessage
      End If
      Set mObj = Nothing
   Else
      Unload Me
   End If
End Sub

Private Sub Label3_Click()
   For mI = 0 To Combo1.UBound
      Combo1(mI).ListIndex = -1
      Text2(mI).Text = ""
   Next
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Label3.BorderStyle = 1
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Label3.BorderStyle = 0
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 0 Then
      KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
   Else
      KeyAscii = fHoraKeyPress(Text1(Index), KeyAscii)
   End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
End Sub

Private Sub sInitForm()
   For mI = 0 To Combo1.UBound
      Combo1(mI).AddItem "OK"
      Combo1(mI).AddItem "KO"
   Next
   Text1(0).Text = Date
   Text1(1).Text = Format(Time, "HH:mm")
End Sub

Private Function fValidar() As Boolean
   fValidar = Fecha_ok(Trim(Text1(0).Text))
   fValidar = Hora_ok(Trim(Text1(1).Text)) And fValidar
End Function
