VERSION 5.00
Begin VB.Form Viol3_frm 
   BackColor       =   &H00B3C1CC&
   Caption         =   "Módulo de ABM de "
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7455
   Begin VB.CommandButton Command1 
      BackColor       =   &H00B3C1CC&
      Caption         =   "Eliminar"
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   10
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
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
      Index           =   3
      Left            =   6240
      TabIndex        =   9
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00B3C1CC&
      Caption         =   "Modificar"
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   8
      Top             =   3720
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2160
      MaxLength       =   30
      TabIndex        =   6
      Top             =   1800
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   3
      Left            =   1800
      MouseIcon       =   "Viol3_frm.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Viol3_frm.frx":0152
      Stretch         =   -1  'True
      ToolTipText     =   "Último"
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   2
      Left            =   1320
      MouseIcon       =   "Viol3_frm.frx":045C
      MousePointer    =   99  'Custom
      Picture         =   "Viol3_frm.frx":05AE
      Stretch         =   -1  'True
      ToolTipText     =   "Siguiente"
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   1
      Left            =   800
      MouseIcon       =   "Viol3_frm.frx":08B8
      MousePointer    =   99  'Custom
      Picture         =   "Viol3_frm.frx":0A0A
      Stretch         =   -1  'True
      ToolTipText     =   "Anterior"
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   0
      Left            =   240
      MouseIcon       =   "Viol3_frm.frx":0D14
      MousePointer    =   99  'Custom
      Picture         =   "Viol3_frm.frx":0E66
      Stretch         =   -1  'True
      ToolTipText     =   "Primero"
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00B3C1CC&
      Caption         =   "Marca"
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
      TabIndex        =   4
      Top             =   2350
      Width           =   540
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00B3C1CC&
      Caption         =   "Descripción"
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
      Left            =   840
      TabIndex        =   3
      Top             =   1850
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00B3C1CC&
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
      Index           =   0
      Left            =   1200
      TabIndex        =   2
      Top             =   1360
      Width           =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   3120
      X2              =   7080
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   3120
      X2              =   7080
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00B3C1CC&
      Caption         =   "ABM de "
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
      Height          =   255
      Left            =   3480
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Viol3_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public pViol2_View As Boolean
Public pViol16_View As Boolean
Dim mObjViol As New clViolaciones
Dim mRec As New ADODB.Recordset
Dim mUltCod As String
Dim mCodMarcaAnt As String
Dim mI As Integer

Private Sub Form_Load()
   Me.Width = 7575
   Me.Height = 4680
   sAlinearForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mObjViol = Nothing
   Set mRec = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
   Dim mFlag As Boolean
     
   Select Case Index
      Case 0 'Nuevo
         If Command1(0).Caption = "Nuevo" Then
            Command1(0).Caption = "Grabar"
            Command1(1).Visible = False
            Command1(2).Visible = False
            Command1(3).Caption = "Volver"
            If Combo1.Visible Then
               Text1(0).Text = mObjViol.fUltCodigo(Trim(mId(Me.Caption, 18, 15)), Right(Combo1.Text, 2))
               Combo1.Enabled = False
            Else
               Text1(0).Text = mObjViol.fUltCodigo(Trim(mId(Me.Caption, 18, 15)), "")
            End If
            Text1(1).Enabled = True
            Text1(1).Text = ""
            fBotonOnOff False
            
         Else 'Grabar NUEVO REGISTRO
            Select Case LCase(Trim(mId(Me.Caption, 18, 15)))
               Case "colores"
                  mFlag = mObjViol.xInsColores(Trim(Text1(0).Text), Trim(Text1(1).Text))
               Case "modelos"
                  mFlag = mObjViol.xInsModelos(Trim(Text1(0).Text), Right(Combo1.Text, 2), Trim(Text1(1).Text))
               Case "marcas"
                  mFlag = mObjViol.xInsMarcas(Trim(Text1(0).Text), Trim(Text1(1).Text))
            End Select
            If mFlag Then
               Command1(0).Caption = "Nuevo"
               Command1(1).Visible = True
               Command1(2).Visible = True
               Command1(3).Caption = "Salir"
               Combo1.Enabled = True
               Text1(1).Enabled = False
               fBotonOnOff True
            End If
         End If
      Case 1  'MODIFICACION
         If Command1(1).Caption = "Modificar" Then
            Command1(0).Visible = False
            Command1(1).Caption = "Grabar"
            Command1(2).Visible = False
            Command1(3).Caption = "Volver"
            Text1(1).Enabled = True
            mCodMarcaAnt = Right(Combo1.Text, 2)
            Combo1.Enabled = False
            fBotonOnOff False
         Else 'Update de la tabla
             Select Case LCase(Trim(mId(Me.Caption, 18, 15)))
                Case "colores"
                   mObjViol.xUpColores Trim(Text1(0).Text), Trim(Text1(1).Text), False
                Case "modelos"
                   mObjViol.xUpModelos Trim(Text1(0).Text), Right(Combo1.Text, 2), Trim(Text1(1).Text), False
                Case "marcas"
                   mObjViol.xUpMarcas Trim(Text1(0).Text), Trim(Text1(1).Text), False
             End Select
            Command1(0).Visible = True
            Command1(1).Caption = "Modificar"
            Command1(2).Visible = True
            Command1(3).Caption = "Salir"
            Text1(1).Enabled = False
            fBotonOnOff True
         End If
      Case 2
         If Text1(0).Text <> "" Then 'Eliminar el Archivo
            If MsgBox("¿Está seguro que desea Eliminar este Código?", vbYesNo, sMessage) = vbYes Then
               Select Case LCase(Trim(mId(Me.Caption, 18, 15)))
                  Case "colores"
                     mObjViol.xUpColores Trim(Text1(0).Text), Trim(Text1(1).Text), True
                  Case "modelos"
                     mObjViol.xUpModelos Trim(Text1(0).Text), Right(Combo1.Text, 2), Trim(Text1(1).Text), True
                  Case "marcas"
                     mObjViol.xUpMarcas Trim(Text1(0).Text), Trim(Text1(1).Text), True
               End Select
               Text1(0).Text = ""
               Text1(1).Text = ""
            End If
         Else
            MsgBox "Debe elegir un Código", vbInformation, "Sistema de Violaciones"
         End If
           
    Case 3
        If Command1(3).Caption = "Salir" Then
           Unload Me
           If pViol2_View Then
               Viol2_frm.Enabled = True
           End If
           If pViol16_View Then
               Viol16_frm.Enabled = True
           Else
              ShowMenu 5, True, False
           End If
           
        Else
           Command1(0).Caption = "Nuevo"
           Command1(1).Caption = "Modificar"
           Command1(2).Caption = "Eliminar"
           Command1(3).Caption = "Salir"
           Command1(0).Visible = True
           Command1(1).Visible = True
           Command1(2).Visible = True
           Command1(3).Visible = True
           Combo1.Enabled = True
           Text1(1).Enabled = False
           mRec.MoveFirst
           If Combo1.Visible Then
              Text1(0).Text = mRec.Fields(1)
              Text1(1).Text = mRec.Fields(2)
           Else
              Text1(0).Text = mRec.Fields(0)
              Text1(1).Text = mRec.Fields(1)
           End If
           fBotonOnOff True
        End If
  End Select
End Sub

Private Sub Combo1_Click()
   If Combo1.ListIndex > -1 Then
      mRec.Close
      Set mRec = mObjViol.oTablaDina("modelos", "where codmarca='" & Right(Combo1.Text, 2) & "' and baja is null order by 3")
      If Not mRec.EOF Then
         Text1(0).Text = mRec.Fields(1)
         Text1(1).Text = mRec.Fields(2)
      End If
   End If
End Sub

Private Sub Image1_Click(Index As Integer)
   Dim Flag As Boolean
   Flag = False
   Select Case Index
      Case 0
         mRec.MoveFirst
         Flag = True
      Case 1
         If Not mRec.BOF Then
            mRec.MovePrevious
            If Not mRec.BOF Then
               Flag = True
            Else
               MsgBox "Se encuentra en el Primer Dato"
            End If
         Else
            MsgBox "Se encuentra en el Primer Dato"
         End If
      Case 2
         If Not mRec.EOF Then
            mRec.MoveNext
            If Not mRec.EOF Then
               Flag = True
            Else
               MsgBox "Se encuentra en el Último Dato"
            End If
         Else
            MsgBox "Se encuentra en el Último Dato"
         End If
      Case 3
         mRec.MoveLast
         Flag = True
   End Select
   If Flag Then
      If Combo1.Visible Then
         Text1(0).Text = mRec.Fields(1)
         Text1(1).Text = mRec.Fields(2)
      Else
         Text1(0).Text = mRec.Fields(0)
         Text1(1).Text = mRec.Fields(1)
      End If
 End If
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image1(Index).BorderStyle = 1
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image1(Index).BorderStyle = 0
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = fAlfaNumKeyPress(KeyAscii)
   KeyAscii = fUcaseKeyPress(KeyAscii)
End Sub

Private Function fBotonOnOff(ByVal pLlave As Boolean)
   For mI = 0 To Image1.UBound
      Image1(mI).Enabled = pLlave
   Next
End Function

Public Sub sInitForm(ByVal pIndex As Integer)
   Select Case pIndex
      Case 0
         Set mRec = mObjViol.oTablaDina("colores", " where baja is null order by 2")
         Me.Caption = Left(Me.Caption, 18) & "Colores"
         Label1.Caption = Left(Label1.Caption, 8) & "Colores"
         Combo1.Visible = False
         Label2(2).Visible = False
      Case 1
         Me.Caption = Left(Me.Caption, 18) & "Modelos"
         Label1.Caption = Left(Label1.Caption, 8) & "Modelos"
         Set mRec = mObjViol.oTabla("marcas", "order by 1")
         sLlenoCbo Viol3_frm.Combo1, mRec, 1, 0
         Set mRec = mObjViol.oTablaDina("modelos", "where codmarca='" & mCodMarcaAnt & "' and baja is null order by 3")
         mCodMarcaAnt = ""
         If Not mRec.EOF Then
            Text1(0).Text = mRec.Fields(1)
            Text1(1).Text = mRec.Fields(2)
         End If
         Combo1.ListIndex = 0
      Case 2
         Me.Caption = Left(Me.Caption, 18) & "Marcas"
         Label1.Caption = Left(Label1.Caption, 8) & "Marcas"
         Combo1.Visible = False
         Label2(2).Visible = False
         Set mRec = mObjViol.oTablaDina("marcas", "where baja is null order by 2")
         If Not mRec.EOF Then
            Text1(0).Text = mRec.Fields(0)
            Text1(1).Text = mRec.Fields(1)
         End If
   End Select
End Sub
