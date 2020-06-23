VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form MEdfrm04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modulo de Solicitudes de Reparaciones (Mant. Edilicio)"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5190
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   14400
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "Solicitud Especial"
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
      Left            =   10920
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   5040
      MaxLength       =   150
      TabIndex        =   5
      Top             =   1560
      Width           =   5655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   4200
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Grabar Solicitud"
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   8040
      TabIndex        =   10
      Top             =   4400
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Volver"
      Height          =   375
      Index           =   2
      Left            =   5640
      TabIndex        =   9
      Top             =   4520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Agregar"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3480
      TabIndex        =   8
      Top             =   4520
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   2640
      MaxLength       =   90
      TabIndex        =   1
      Top             =   840
      Width           =   8055
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   2480
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   2990
      _Version        =   327680
      Cols            =   8
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Observaciones"
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
      Left            =   5280
      TabIndex        =   17
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "D.Estim."
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
      Left            =   4090
      TabIndex        =   16
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Rubro"
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
      Left            =   1920
      TabIndex        =   15
      Top             =   1320
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Mantenim."
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
      Left            =   240
      TabIndex        =   14
      Top             =   1320
      Width           =   1320
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   120
      Y1              =   480
      Y2              =   2240
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   14175
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Lugar"
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
      TabIndex        =   13
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Registro de solicitudes de servicios (Mant. Edilicio)"
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
      Left            =   2040
      TabIndex        =   12
      Top             =   45
      Width           =   6105
   End
   Begin VB.Line Line8 
      X1              =   14175
      X2              =   14175
      Y1              =   480
      Y2              =   2240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descripcion del problema"
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
      Left            =   2880
      TabIndex        =   11
      Top             =   600
      Width           =   2160
   End
   Begin VB.Line Line9 
      X1              =   120
      X2              =   14175
      Y1              =   2240
      Y2              =   2240
   End
End
Attribute VB_Name = "MEdfrm04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObj As New clMantEd
'Dim mObjPea As New clPeaje
Dim mRec As New ADODB.Recordset
Dim mObjLuser As New clLogUser

Private Sub Command1_Click(Index As Integer)
Dim mI As Integer
Dim mNroParte As String
Dim mSolicEspecial
Dim mIntSolicEspecial
If Index = 0 Then
   If fValida Then
   
   
      If Check1.Value Then
        mSolicEspecial = "Si"
      Else
        mSolicEspecial = "No"
      End If
   
   
      MSFlexGrid1.AddItem vbTab & Combo1(0).Text & vbTab & Text1(0).Text & vbTab & Combo1(1).Text & vbTab & Combo1(2).Text & vbTab & Text1(1).Text & vbTab & Text1(2).Text & vbTab & mSolicEspecial
      Command1(1).Enabled = True
      If MSFlexGrid1.TextMatrix(1, 1) = "" Then
         MSFlexGrid1.RemoveItem 1
      End If
      'Limpio los textBoxs
      For mI = 0 To Text1.Count - 1
         Text1(mI).Text = ""
      Next
      'Limpio los comboBoxs
      For mI = 0 To Combo1.Count - 1
         Combo1(mI).ListIndex = -1
      Next
      'Limpio checkbox
      Check1.Value = False
      
   End If
Else
   If Index = 1 Then
      mNroParte = mObj.ObtMaxParte
      For mI = 1 To MSFlexGrid1.Rows - 1 'mp20161221
      
         If MSFlexGrid1.TextMatrix(mI, 7) = "Si" Then
            mIntSolicEspecial = 1
         Else
            mIntSolicEspecial = 0
         End If
         mObj.InsRegistros mI + mNroParte, Format(Now, "yyyy-mm-dd hh:mm:ss"), MSFlexGrid1.TextMatrix(mI, 1), MSFlexGrid1.TextMatrix(mI, 2), "Media", "MantEdil", "P", Trim(Right(MDI.mUser, 20)), "M", MSFlexGrid1.TextMatrix(mI, 3), MSFlexGrid1.TextMatrix(mI, 4), MSFlexGrid1.TextMatrix(mI, 5), MSFlexGrid1.TextMatrix(mI, 6), mIntSolicEspecial
         
         
      Next
   End If
   Unload Me
End If
End Sub

Private Sub Form_Load()
Dim mI As Integer
Dim mTramo As String

MEdfrm04.Top = 100
MEdfrm04.Left = (MDI.Width - MEdfrm04.Width) / 2
Check1.RightToLeft = True


'Veo que tramos debe mostrar el combo de Edificios
'mTramo = ""
'Select Case Left(mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 3), InStr(1, mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 3), "@") - 1)
'   Case "marasc1", "realasc", "suptigre", "testigre", "mpacheco"
'      mTramo = "Z1"
'   Case "supcampana", "supcampanadecalado", "tescampana"
'      mTramo = "Z2"
'   Case "suppilar", "suppilardec", "tespilar"
'      mTramo = "Z3"
'   Case "197desc1", "bayre", "belgrano", "buenayredesc", "r202", "ruta197desc", "sup202a"
'      mTramo = "Z4"
'End Select
'


'Set mRec = mObj.oEjecutarSelect("SELECT * FROM Edificios WHERE FechaBaja IS NULL " & IIf(mTramo <> "", " And Tramo = '" & mTramo & "'", ""))
Set mRec = mObj.oEjecutarSelect("SELECT * FROM Edificios WHERE FechaBaja IS NULL ORDER BY ZonaMantEdil, Descripcion")

If Not mRec.EOF Then
   Do While Not mRec.EOF
      'Combo1(0).AddItem mRec!Tramo & " - " & mRec!descripcion
      Combo1(0).AddItem mRec!ZonaMantEdil & " - " & mRec!descripcion
      mRec.MoveNext
   Loop
End If
mRec.Close
Combo1(1).AddItem "Preventivo"
Combo1(1).AddItem "Predictivo"
Combo1(1).AddItem "Correctivo"

Set mRec = mObj.oEjecutarSelect("SELECT * FROM Rubros WHERE FechaBaja IS NULL")
If Not mRec.EOF Then
   Do While Not mRec.EOF
      Combo1(2).AddItem mRec!Codigo & " - " & mRec!descripcion
      mRec.MoveNext
   Loop
End If
mRec.Close

MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 1800
MSFlexGrid1.ColWidth(2) = 4400
MSFlexGrid1.ColWidth(3) = 900
MSFlexGrid1.ColWidth(4) = 1800
MSFlexGrid1.ColWidth(5) = 500
MSFlexGrid1.ColWidth(6) = 2900
'MSFlexGrid1.ColWidth(7) = 1500
MSFlexGrid1.ColWidth(7) = 0
MSFlexGrid1.TextMatrix(0, 1) = "Edificio"
MSFlexGrid1.TextMatrix(0, 2) = "Descripcion del Problema"
MSFlexGrid1.TextMatrix(0, 3) = "Tipo Mant."
MSFlexGrid1.TextMatrix(0, 4) = "Rubro"
MSFlexGrid1.TextMatrix(0, 5) = "Estim"
MSFlexGrid1.TextMatrix(0, 6) = "Observaciones"
'MSFlexGrid1.TextMatrix(0, 7) = "Solicitud Especial"
MSFlexGrid1.TextMatrix(0, 7) = ""
For mI = 0 To MSFlexGrid1.Cols - 1
   MSFlexGrid1.ColAlignment(mI) = 0
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mObj = Nothing
Set mRec = Nothing
Set mObjLuser = Nothing
ShowMenu 20, True, False
End Sub

Private Function fValida() As Boolean
Dim mRet As Boolean
Dim mI As Integer
Dim mRepe As Boolean
mRepe = False
mRet = (Text1(0).Text <> "" And Text1(1).Text <> "")
If mRet Then
   For mI = 1 To Combo1.UBound
      If mRet Then
         mRet = (Combo1(mI).Text <> "")
      End If
   Next
End If
If mRet Then
   If MSFlexGrid1.Rows > 1 Then
      For mI = 1 To MSFlexGrid1.Rows - 1
         If MSFlexGrid1.TextMatrix(mI, 1) = Combo1(0).Text And MSFlexGrid1.TextMatrix(mI, 2) = Text1(0).Text And mRet Then
            mRet = False
            mRepe = True
         End If
      Next
   End If
End If
If Not mRet Then
   If mRepe Then
      MsgBox "Ya está ingresada dicha solicitud", vbCritical, "Atención"
   Else
      MsgBox "Verifique que todos los datos estén ingresados correctamente", vbCritical, "Atención"
   End If
End If
fValida = mRet
End Function

Private Sub MSFlexGrid1_DblClick()
Dim mI As Integer
Dim mJ As Integer
If MSFlexGrid1.Row > 0 And MSFlexGrid1.TextMatrix(1, 1) <> "" And MSFlexGrid1.Col = 1 Then
   If MsgBox("¿Está Seguro de Eliminar este Registro?", vbYesNo, sMessage) = vbYes Then
      If MSFlexGrid1.Rows > 2 Then
         For mI = MSFlexGrid1.Row To MSFlexGrid1.Rows - 2
            For mJ = 1 To MSFlexGrid1.Cols - 1
               MSFlexGrid1.TextMatrix(mI, mJ) = MSFlexGrid1.TextMatrix(mI + 1, mJ)
            Next
         Next
         MSFlexGrid1.RemoveItem (MSFlexGrid1.Rows - 1)
      Else
         MSFlexGrid1.AddItem ""
         MSFlexGrid1.RemoveItem 1
      End If
   End If
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
   Case 1
      KeyAscii = fNumeroKeyPress(KeyAscii)
End Select
End Sub
