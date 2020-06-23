VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form RNov8_frm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seguimiento de Novedades"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   7455
      Left            =   70
      TabIndex        =   6
      Top             =   720
      Width           =   11700
      Begin VB.Timer Timer2 
         Left            =   11040
         Top             =   240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
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
         Left            =   9720
         TabIndex        =   9
         Top             =   300
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   11033
         _Version        =   327680
         Rows            =   1
         Cols            =   7
         FixedCols       =   0
         FillStyle       =   1
         ScrollBars      =   2
         MergeCells      =   2
      End
      Begin VB.Label Label4 
         Caption         =   "Seguimiento de Novedades"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   5400
      TabIndex        =   2
      Top             =   0
      Width           =   6380
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   4320
         Top             =   120
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4920
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   3
         Top             =   330
         Width           =   60
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   70
      TabIndex        =   0
      Top             =   0
      Width           =   5265
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sistema de Registro de Novedades"
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
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   4245
      End
   End
End
Attribute VB_Name = "RNov8_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObjRN As New clRNov
Dim mRec As New ADODB.Recordset
Public xUltFecha As Date

Private Sub Form_Load()
   sInitForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mObjRN = Nothing
   Set mRec = Nothing
End Sub

Private Sub Command1_Click()
  Unload RNov8_frm
  ShowMenu 1, True, False
End Sub

Private Sub Timer1_Timer()
   Label2.Caption = Right(Now, 8)
   Label2.ToolTipText = Left(Now, 10)
End Sub

Private Sub MSFlexGrid1_DblClick()
Dim Texto As String
Dim mCodigo As String
Dim mKm As Double
Dim mFecha As Date
   
   If MSFlexGrid1.Col = 4 Then
      If MSFlexGrid1.Text = "SI" Then
         MSFlexGrid1.Col = 0
         mCodigo = Trim(MSFlexGrid1.Text)
         MSFlexGrid1.Col = 1
         mFecha = MSFlexGrid1.Text
         Set mRec = mObjRN.oMovilesXCodigo(mCodigo, mFecha)
         Texto = "Móviles Asignados:   " & NVL(mRec!Mov1, "") & " - " & NVL(mRec!Mov2, "") & " - " & NVL(mRec!Mov3, "")
         mRec.Close
         MsgBox Texto, vbInformation, sMessage
      Else
         If MSFlexGrid1.CellPicture <> 0 Then
            MsgBox "Es una Demora.", vbInformation, sMessage
         End If
      End If
   End If
End Sub

Private Sub Timer2_Timer()
Dim mI As Integer
Dim Cod As String
Dim mMov As String
Dim xSent As String
Dim xCodRamal As String
Dim xFecha As String
Dim mObj As New clRNov

   Set mRec = mObjRN.oTabla("novedades2", " where fecha > '" & Format(xUltFecha, "yyyy/mm/dd hh:mm:ss") & "'")
   
   'Set mRec = mObjRN.oTabla("novedades2", "where fecha > DATE_ADD(now(),INTERVAL -1 DAY) order by fecha")
   If Not mRec.EOF Then
      Do While Not mRec.EOF
          mMov = ""
          Cod = NVL(mRec!Codigo, "")
          xSent = NVL(mRec!Sent, "")
          If VarType(mRec!Sent) = 1 Or mRec!Sent = 0 Then
            xSent = ""
            xCodRamal = ""
          Else
            xSent = Left(mObj.sTablaDescr("sentidos", "codigo=" & mRec!Sent, 1), 2)
            xCodRamal = Mid(mObj.sTablaDescr("ramales", "codigo=" & mRec!codramal, 2), 2, 2)
          End If
      
          
          
          If mRec!CodNov = "A" Or mRec!CodNov = "C" Or mRec!CodNov = "E" Then
             mMov = "SI"
          End If
          sCargar Cod, mRec!Fecha, mRec!Km, xSent, xCodRamal, mMov, mRec!Descripcion, MSFlexGrid1, True
          If mRec!CodNov = "D" Then
             MSFlexGrid1.Col = 5
             If MSFlexGrid1.CellBackColor = &HFFFFFF Then
                Set MSFlexGrid1.CellPicture = LoadPicture(App.Path & "\RegNovedades\Image\Reloj.bmp")
             Else
                Set MSFlexGrid1.CellPicture = LoadPicture(App.Path & "\RegNovedades\Image\Reloja.bmp")
             End If
          End If
          mRec.MoveNext
       Loop
       mRec.MovePrevious
       xFecha = mRec!Fecha
       RNov8_frm.Refresh
       MsgBox "Se Importaron Nuevos Datos", vbInformation, sMessage
   End If
   mRec.Close
   Set mObj = Nothing
End Sub

Private Sub sInitForm()
Dim mI As Integer
Dim mCont As Integer
Dim Cod As String
Dim xSent As String
Dim xCodRamal As String
Dim mMov As String
Dim mObj As New clRNov

   sMsgEspere Me, "Cargando datos... aguarde un momento.", True
   Timer2.Interval = 60000
   sAlinearForm Me
   Label2.Caption = Right(Now, 8)
   Label2.ToolTipText = Left(Now, 10)
   Label3(1).Caption = Trim(Left(MDI.mUser, 35))
   With MSFlexGrid1
      .ColWidth(0) = 900
      .ColWidth(1) = 1600
      .ColWidth(2) = 600
      .ColWidth(3) = 400
      .ColWidth(4) = 400
      .ColWidth(5) = 450
      .ColWidth(6) = 6580
      .TextMatrix(0, 0) = "Código"
      .TextMatrix(0, 1) = "Fecha/Hora"
      .TextMatrix(0, 2) = "Km"
      .TextMatrix(0, 3) = "Sen"
      .TextMatrix(0, 4) = "Ram"
      .TextMatrix(0, 5) = "Mov"
      .TextMatrix(0, 6) = "Novedad"
      sSetFlexColOrder RNov8_frm.MSFlexGrid1, 0
      .MergeCol(0) = True
   End With
   Set mRec = mObjRN.oTabla("novedades2", "where fecha > DATE_ADD(now(),INTERVAL -1 DAY) order by fecha")
   mI = 1
   Do While Not mRec.EOF
      mMov = ""
      MSFlexGrid1.Font = "Arial"
      Cod = NVL(mRec!Codigo, "")
      xSent = NVL(mRec!Sent, "")
      If mRec!CodNov = "A" Or mRec!CodNov = "C" Or mRec!CodNov = "E" Then
         mMov = "SI"
      End If
      If VarType(mRec!Sent) = 1 Or mRec!Sent = 0 Then
         xSent = ""
         xCodRamal = ""
      Else
         xSent = Left(mObj.sTablaDescr("sentidos", "codigo=" & mRec!Sent, 1), 2)
         xCodRamal = Mid(mObj.sTablaDescr("ramales", "codigo=" & mRec!codramal, 2), 2, 2)
      End If
      
      
      If mI < 3 Then
         sCargar Cod, mRec!Fecha, mRec!Km, xSent, xCodRamal, mMov, mRec!Descripcion, MSFlexGrid1, True
      Else
         sCargar Cod, mRec!Fecha, mRec!Km, xSent, xCodRamal, mMov, mRec!Descripcion, MSFlexGrid1, False
      End If
      mI = mI + 1
      If mRec!CodNov = "D" Then
         MSFlexGrid1.Col = 5
         If MSFlexGrid1.CellBackColor = &HFFFFFF Then
            Set MSFlexGrid1.CellPicture = LoadPicture(App.Path & "\RegNovedades\Image\Reloj.bmp")
         Else
             Set MSFlexGrid1.CellPicture = LoadPicture(App.Path & "\RegNovedades\Image\Reloja.bmp")
         End If
      End If
      xUltFecha = mRec!Fecha
      mRec.MoveNext
   Loop
   mRec.Close
   sMsgEspere Me, "", False
   RNov8_frm.Refresh
   Set mObj = Nothing
End Sub

Public Function sCargar(ByVal xCod As String, xFecha As Date, ByVal xKm As Double, ByVal xSent As String, xRamal As String, ByVal xMov As String, ByVal xNov As String, xFlex As Object, ByVal mBold As Boolean)
Dim mColor As Boolean
Dim mFlag As Boolean
Dim mI As Integer
Dim mRow As Integer
Dim mColour As Double
Dim Kolor As Double
 
   mColor = False
   mFlag = False
   If xFlex.Rows = 1 Then 'Es el primero
      mColour = &HC0FFFF
      'xFlex.AddItem xCod & vbTab & xFecha & vbTab & Format(xKm, "#0.00") & vbTab & xSent & vbTab & xRamal & vbTab & xMov & vbTab & xNov
      xFlex.AddItem xCod & vbTab & xFecha & vbTab & xKm & vbTab & xSent & vbTab & xRamal & vbTab & xMov & vbTab & xNov
      xFlex.Row = 1
      For mI = 0 To 6
         xFlex.Col = mI
         xFlex.CellBackColor = mColour
      Next
   Else
      xFlex.Col = 0
      For mI = 1 To xFlex.Rows - 1
         xFlex.Row = mI
         If xFlex.Text = xCod Then
            mRow = mI
            mFlag = True
         End If
      Next
      If mFlag And xCod <> "" Then
         xFlex.Row = mRow
         xFlex.AddItem xCod & vbTab & xFecha & vbTab & Format(xKm, "#0.00") & vbTab & xSent & vbTab & xRamal & vbTab & xMov & vbTab & xNov, mRow + 1
         mColour = xFlex.CellBackColor
         xFlex.MergeRow(mRow + 1) = True
         xFlex.Col = 0
         xFlex.Row = mRow + 1
         For mI = 0 To 6
            xFlex.Col = mI
            xFlex.CellBackColor = mColour
         Next
      Else
         xFlex.Row = 1
         xFlex.Col = 1
         mColour = &HC0FFFF
         Kolor = xFlex.CellBackColor
         If Kolor = mColour Then
            mColour = &HFFFFFF
         End If
         xFlex.Col = 0
         xFlex.AddItem xCod & vbTab & xFecha & vbTab & Format(xKm, "#0.00") & vbTab & xSent & vbTab & xRamal & vbTab & xMov & vbTab & xNov, 1
         xFlex.Row = 1
         For mI = 0 To 6
            xFlex.Col = mI
            xFlex.CellBackColor = mColour
         Next
      End If
      End If
      xFlex.Col = 0
      xFlex.CellAlignment = 4
      xFlex.Col = 3
      xFlex.CellAlignment = 4
      xFlex.Col = 4
      xFlex.CellAlignment = 4
      If mBold Then
         For mI = 1 To 6
            xFlex.Col = mI
            xFlex.CellFontBold = 1
            xFlex.CellForeColor = &HFF0000
         Next
   End If
End Function
