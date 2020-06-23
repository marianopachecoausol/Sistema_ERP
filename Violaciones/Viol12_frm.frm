VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Viol12_frm 
   BackColor       =   &H00CECECE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulo ABM de Código Postales."
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   ControlBox      =   0   'False
   ForeColor       =   &H00400000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   7125
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   8
      Top             =   960
      Width           =   1335
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
      Height          =   495
      Index           =   0
      Left            =   5640
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid mFlex 
      Height          =   4935
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8705
      _Version        =   327680
      FixedCols       =   0
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "Localidad"
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
      TabIndex        =   2
      Top             =   645
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "Provincia"
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
      TabIndex        =   1
      Top             =   1005
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "Código Postal"
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
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   1185
   End
End
Attribute VB_Name = "Viol12_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObjViol As New clViolaciones
Dim mRec As New ADODB.Recordset
Dim mI As Integer
Public mSistemaActivo As Integer

Private Sub Form_Load()
Me.Width = 7215
Me.Height = 6945
sAlinearForm Me
Set mRec = mObjViol.oTabla("provincias", "order by 1")
Do While Not mRec.EOF
   Combo1.AddItem mRec.Fields(0) & " - " & mRec.Fields(1)
   mRec.MoveNext
Loop
mRec.Close
sTituloFlex
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mObjViol = Nothing
Set mRec = Nothing
If mSistemaActivo = 25 Then
   SEmpr1_frm.Enabled = True
Else
   ShowMenu 5, True, False
End If
End Sub

Private Sub Combo1_Click()
If Text1(0).Enabled = False And Text1(1).Enabled = False Then
   If Combo1.ListIndex > -1 Then
      For mI = mFlex.Rows - 1 To 2 Step -1
         mFlex.RemoveItem mI
      Next
      mFlex.Clear
      Set mRec = mObjViol.oCP_Pcia(Left(Combo1.Text, 2))
      Do While Not mRec.EOF
         mFlex.AddItem mRec.Fields(0) & vbTab & mRec.Fields(2)
         mRec.MoveNext
      Loop
      mRec.Close
      If mFlex.Rows > 3 Then
         mFlex.RemoveItem 1
      End If
      sTituloFlex
      sSetFlex2Colors Viol12_frm.mFlex, &HE1FFFD, &HF3FEF1
   End If
End If
End Sub

Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
   If Command1(0).Caption = "Nuevo" Then
      Text1(0).Enabled = True
      Text1(1).Enabled = True
      mFlex.Enabled = False
      Combo1.Enabled = True
      Command1(0).Caption = "Grabar"
      Command1(1).Caption = "Volver"
   Else
      If Text1(0).Text <> "" And Text1(1).Text <> "" And Combo1.ListIndex > -1 Then
         If Command1(0).Caption = "Grabar" Then
            If Not mObjViol.xInsertCP(Trim(Text1(0).Text), Left(Combo1.Text, 2), Trim(Text1(1).Text)) Then
               MsgBox "El código que quiere ingresar Ya Existe", vbCritical, sMessage & "Atención!"
            End If
         Else
            If Not mObjViol.xUpdateCP(Trim(Text1(1).Text), Left(Combo1.Text, 2), Trim(Text1(0).Text), Combo1.Tag) Then
               MsgBox "Error en la Actualización", vbCritical, sMessage & "Atención!"
            End If
         End If
         sVolver
      Else
         MsgBox "Existe un Error!", vbCritical, sMessage & "Atención!"
      End If
   End If
Else
   If Command1(1).Caption = "Volver" Then
      sVolver
   Else
      Unload Me
   End If
End If
End Sub

Private Sub mFlex_DblClick()
If mFlex.TextMatrix(mFlex.Row, 0) <> "" Then
   Text1(0).Text = mFlex.TextMatrix(mFlex.Row, 0)
   Text1(1).Text = mFlex.TextMatrix(mFlex.Row, 1)
   Text1(1).Enabled = True
   Command1(0).Caption = "Modificar"
   Command1(1).Caption = "Volver"
   mFlex.Enabled = False
   mFlex.Col = 0
   mFlex.ColSel = 1
   Combo1.Enabled = True
   Combo1.Tag = Left(Combo1.Text, 2)
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
KeyAscii = fUcaseKeyPress(KeyAscii)
End Sub

Private Sub sTituloFlex()
With mFlex
   .ColWidth(0) = 1200
   .ColWidth(1) = 5500
   .Row = 0
   .Font = "Arial"
   .Col = 0
   .CellFontBold = True
   .FixedAlignment(0) = 4
   .Text = "Cód. Postal"
   .Col = 1
   .CellFontBold = True
   .FixedAlignment(1) = 4
   .Text = "Descripción"
End With
End Sub

Private Sub sVolver()
Text1(0).Text = ""
Text1(1).Text = ""
Text1(0).Enabled = False
Text1(1).Enabled = False
mFlex.Enabled = True
Combo1.Enabled = True
Command1(0).Caption = "Nuevo"
Command1(1).Caption = "Salir"
Combo1.Tag = ""
End Sub
