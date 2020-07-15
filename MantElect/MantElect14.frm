VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MantElect14 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nuevo Comunicado"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16965
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   16965
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Datos del Comunicado"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   16680
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   12
         Top             =   450
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   5760
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   465
         Width           =   2895
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   4800
         TabIndex        =   13
         Top             =   510
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Comunicado:"
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
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   510
         Width           =   1095
      End
   End
   Begin VB.CommandButton CommandProd 
      Height          =   375
      Left            =   8280
      Picture         =   "MantElect14.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3675
      Width           =   330
   End
   Begin VB.Frame Frame1 
      Caption         =   "Generar partes"
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
      Height          =   2130
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   16680
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   15120
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   8880
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   480
         Width           =   3975
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   5760
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1800
         MaxLength       =   44
         TabIndex        =   17
         Top             =   1440
         Width           =   4935
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   8880
         MaxLength       =   44
         TabIndex        =   6
         Top             =   960
         Width           =   4635
      End
      Begin VB.Label Label10 
         Caption         =   "Prioridad:"
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
         Height          =   255
         Left            =   14160
         TabIndex        =   27
         Top             =   540
         Width           =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   13840
         X2              =   13840
         Y1              =   120
         Y2              =   2160
      End
      Begin VB.Label Label9 
         Caption         =   "Problema:"
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
         Height          =   255
         Left            =   7725
         TabIndex        =   24
         Top             =   540
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   6960
         X2              =   6960
         Y1              =   120
         Y2              =   2160
      End
      Begin VB.Label Label8 
         Caption         =   "Progresiva:"
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
         Height          =   255
         Left            =   4680
         TabIndex        =   22
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label7 
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3040
         TabIndex        =   20
         Top             =   1020
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Otro Activo:"
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
         Height          =   255
         Left            =   520
         TabIndex        =   19
         Top             =   1500
         Width           =   1140
      End
      Begin VB.Label Label5 
         Caption         =   "Otro Problema:"
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
         Height          =   255
         Left            =   7320
         TabIndex        =   18
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Acceso:"
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
         Height          =   255
         Left            =   840
         TabIndex        =   15
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo Activo:"
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
         Height          =   255
         Left            =   505
         TabIndex        =   7
         Top             =   540
         Width           =   1080
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Confirmar"
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   4
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Index           =   1
      Left            =   9480
      TabIndex        =   3
      Top             =   8880
      Width           =   1815
   End
   Begin VB.Frame Frame10 
      Caption         =   "Partes generados"
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
      Height          =   4560
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   16680
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   278
         Left            =   13680
         TabIndex        =   1
         Top             =   7555
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid FlexPartes 
         Height          =   3960
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   16440
         _ExtentX        =   28998
         _ExtentY        =   6985
         _Version        =   327680
         Cols            =   6
      End
   End
End
Attribute VB_Name = "MantElect14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mObj As New clMantElect
Dim mRec As New ADODB.Recordset

Dim filaAnt As Integer
Dim columnAnt As Integer

Dim mvMat_CodProd_Orig() As String
Dim mvMat_CodUbic_Orig() As String
Dim mvMat_Cantidad_Orig() As Double
Dim mvMat_CantidadBD_Orig() As Double


Dim cboRamalListIndex As Integer

Private Sub Combo1_Click()
   Dim mi As Integer
   If cboRamalListIndex <> Combo1.ListIndex Then
      
      If cboRamalListIndex <> -1 And FlexPartes.Rows > 2 Then
         If MsgBox("Si selecciona otro Ramal se perderán los datos cargados hasta el momento. ¿ Desea continuar ? ", vbYesNo, "Cambio de Ramal") = vbYes Then
            limpiarFormularioParcial
            'Elimino los registros grilla inferior
            For mi = FlexPartes.Rows To 3 Step -1
               FlexPartes.RemoveItem mi
            Next
         Else
            Combo1.ListIndex = cboRamalListIndex
         End If
         cboRamalListIndex = Combo1.ListIndex
      
      
      Else
         limpiarFormularioParcial
         cboRamalListIndex = Combo1.ListIndex
      End If
   End If
End Sub

Private Sub Combo2_Click()
   'TODO: Validar que comboRamal tenga algo elegido.
   If Combo1.Text <> "" Then
      sLlenoComboAcceso
      sLlenoComboProblema
   Else
      MsgBox "Debe seleccionar un Ramal", vbExclamation, "Atención !!!"
      sLlenoTipoActivo
   End If

End Sub

Private Sub sLlenoComboAcceso()
   Dim mCodRamal As String
   Dim mCodTipoActivo As String
   'Dim mObj As New clInven
   Dim mRec1 As New ADODB.Recordset
   
   
   mCodRamal = Right(Combo1.Text, 1)
   mCodTipoActivo = Right(Combo2.Text, 2)
   Combo3.Clear
   Combo4.Clear
   Combo5.Clear

   Set mRec1 = mObj.oEjecutarSelect("SELECT DISTINCT Acceso FROM COM_Activos " & _
                                    " WHERE CodRamal = '" & mCodRamal & "'" & _
                                    " AND CodTipoActivo = '" & mCodTipoActivo & "' ORDER BY Acceso")
   
      
  If Not mRec1.EOF Then
     Combo3.Enabled = True
      Text3.Text = ""
      Text3.Enabled = False
      
      
      Do While Not mRec1.EOF
         Combo3.AddItem "" & mRec1!Acceso
         mRec1.MoveNext
      Loop
   Else
      Combo3.Enabled = False
      Combo4.Enabled = False
      Combo5.Enabled = False
      Text3.Enabled = True
      Text3.SetFocus
   End If
   
   
   
   
   mRec1.Close
   'Set mObj = Nothing
   Set mRec1 = Nothing
End Sub

Private Sub sLlenoComboProblema()

   Dim mCodTipoActivo As String
   Dim mRec1 As New ADODB.Recordset
   
   mCodTipoActivo = Right(Combo2.Text, 2)
   Combo6.Clear

   Set mRec1 = mObj.oEjecutarSelect("SELECT * FROM COM_TiposActivo_Problemas " & _
                                    " WHERE CodTipoActivo = '" & mCodTipoActivo & "'" & _
                                    " UNION" & _
                                    " SELECT * FROM COM_TiposActivo_Problemas" & _
                                    " WHERE CodTipoActivo Is Null " & _
                                    " ORDER BY Codigo")
   
   Do While Not mRec1.EOF
      Combo6.AddItem "" & mRec1!descripcion & Space(100) & mRec1!Codigo
      mRec1.MoveNext
   Loop
   
   mRec1.Close
   'Set mObj = Nothing
   Set mRec1 = Nothing
End Sub

Private Sub Combo3_Click()
   sLlenoComboKm
End Sub

Private Sub sLlenoComboKm()
   Dim mCodRamal As String
   Dim mCodTipoAcceso As String
   Dim mCodAcceso As String
   
   Dim mRec1 As New ADODB.Recordset
   
   mCodRamal = Right(Combo1.Text, 1)
   mCodTipoAcceso = Right(Combo2.Text, 2)
   mCodAcceso = Trim(Combo3.Text)
   
   
   Combo4.Clear
   Combo4.Enabled = True
   Combo5.Clear

   Set mRec1 = mObj.oEjecutarSelect("SELECT DISTINCT Km FROM COM_Activos " & _
                                    " WHERE CodRamal = '" & mCodRamal & "'" & _
                                    " AND CodTipoActivo = '" & mCodTipoAcceso & "'" & _
                                    " AND Acceso = '" & mCodAcceso & "'" & _
                                    " ORDER BY Km")
   
   Do While Not mRec1.EOF
     Combo4.AddItem "" & mRec1!km
     mRec1.MoveNext
   Loop
   mRec1.Close
   'Set mObj = Nothing
   Set mRec1 = Nothing
End Sub

Private Sub Combo4_Click()
   sLlenoComboProgresiva
End Sub

Private Sub sLlenoComboProgresiva()
   Dim mCodRamal As String
   Dim mCodTipoAcceso As String
   Dim mCodAcceso As String
   Dim mKm As String
   
   Dim mRec1 As New ADODB.Recordset
   
   mCodRamal = Right(Combo1.Text, 1)
   mCodTipoAcceso = Right(Combo2.Text, 2)
   mCodAcceso = Trim(Combo3.Text)
   mKm = Trim(Combo4.Text)
   
   Combo5.Clear
   Combo5.Enabled = True
   
   Set mRec1 = mObj.oEjecutarSelect("SELECT DISTINCT Progresiva, Codigo FROM COM_Activos " & _
                                    " WHERE CodRamal = '" & mCodRamal & "'" & _
                                    " AND CodTipoActivo = '" & mCodTipoAcceso & "'" & _
                                    " AND Acceso = '" & mCodAcceso & "'" & _
                                    " AND Km = '" & mKm & "'" & _
                                    " ORDER BY Progresiva")
   
   Do While Not mRec1.EOF
     Combo5.AddItem "" & mRec1!Progresiva & Space(100) & mRec1!Codigo
     mRec1.MoveNext
   Loop
   mRec1.Close
   'Set mObj = Nothing
   Set mRec1 = Nothing
End Sub

Private Sub Combo6_Click()
   Dim mRec1 As New ADODB.Recordset
   Dim mCodTipoActivo_Problema As String

   Text4.Text = ""
   mCodTipoActivo_Problema = Right(Combo6.Text, 2)

   Set mRec1 = mObj.oEjecutarSelect("SELECT Codigo, Descripcion, IFNULL(CodTipoActivo,'') AS CodTipoActivo " & _
                                    " FROM COM_TiposActivo_Problemas " & _
                                    " WHERE Codigo = '" & mCodTipoActivo_Problema & "'")
                                    

   If mRec1!CodTipoActivo = "" Then
      Text4.Enabled = True
      Text4.SetFocus
   Else
      Text4.Enabled = False
   End If


   mRec1.Close
   'Set mObj = Nothing
   Set mRec1 = Nothing
End Sub


Private Sub Command2_Click(Index As Integer)
   Dim vParte_CodActivo() As String
   Dim vParte_DescActivo() As String
   Dim vParte_CodProblema() As String
   Dim vParte_DescProblema() As String
   Dim vParte_Prioridad() As String
   Dim vParte_NroParte() As Integer
   Dim mNroComunicado As String
   Dim mCodRamal As String
   
   If Index = 0 Then
      If fValidaComunicado() Then
         preparaArrayPartes vParte_NroParte(), vParte_CodActivo(), vParte_DescActivo(), vParte_CodProblema(), vParte_DescProblema(), vParte_Prioridad()
         
         mNroComunicado = Trim(Text1.Text)
         mCodRamal = Right(Combo1.Text, 1)
         
         mObj.xGenerarComunicado mNroComunicado, mCodRamal, vParte_NroParte(), vParte_CodActivo(), vParte_DescActivo(), vParte_CodProblema(), vParte_DescProblema(), vParte_Prioridad(), Trim(Right(MDI.mUser, 15))
      
         limpiarFormularioTotal
      End If
   Else
      Unload Me
   End If
End Sub

Private Sub preparaArrayPartes(ByRef pvParte_NroParte() As Integer, ByRef pvParte_CodActivo() As String, ByRef pvParte_DescActivo() As String, ByRef pvParte_CodProblema() As String, ByRef pvParte_DescProblema() As String, ByRef pvParte_Prioridad() As String)
   Dim mj As Integer
   Dim cantPartes As Integer

   cantPartes = FlexPartes.Rows - 2
   If cantPartes > 0 Then
      
      ReDim pvParte_CodActivo(0 To cantPartes - 1) As String
      ReDim pvParte_DescActivo(0 To cantPartes - 1) As String
      ReDim pvParte_CodProblema(0 To cantPartes - 1) As String
      ReDim pvParte_DescProblema(0 To cantPartes - 1) As String
      ReDim pvParte_Prioridad(0 To cantPartes - 1) As String
      ReDim pvParte_NroParte(0 To cantPartes - 1) As Integer
      
      For mj = 2 To FlexPartes.Rows - 1
        pvParte_CodActivo(mj - 2) = FlexPartes.TextMatrix(mj, 4)
        pvParte_DescActivo(mj - 2) = FlexPartes.TextMatrix(mj, 1)
        pvParte_CodProblema(mj - 2) = FlexPartes.TextMatrix(mj, 5)
        pvParte_DescProblema(mj - 2) = FlexPartes.TextMatrix(mj, 2)
        pvParte_Prioridad(mj - 2) = FlexPartes.TextMatrix(mj, 3)
        pvParte_NroParte(mj - 2) = 0
      Next
   End If
End Sub


Private Function fValidaComunicado() As Boolean
   Dim mRet As Boolean
   Dim mMensajeError As String
   Dim mRec1 As New ADODB.Recordset
   
   mRet = True
   
   If Trim(Text1.Text) = "" Then
      mRet = False
      mMensajeError = "Debe completar el campo 'Comunicado'"
   End If
   
   If mRet Then
      Set mRec1 = mObj.oEjecutarSelect("SELECT * FROM COM_Comunicados_H WHERE NroComunicado = '" & Trim(Text1.Text) & "';")
      If Not mRec1.EOF Then
         mRet = False
         mMensajeError = "Ya existe un comunicado con ese número."
      End If
      mRec1.Close
   End If
   
   If mRet Then
      If Trim(Combo1.Text) = "" Then
         mRet = False
         mMensajeError = "Debe seleccionar un Ramal"
      End If
   End If
   
   If mRet Then
      If FlexPartes.Rows <= 2 Then
         mRet = False
         mMensajeError = "Al menos se debe crear un Parte"
      End If
   End If
   
   Set mRec1 = Nothing
   
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If

   fValidaComunicado = mRet
End Function


Private Sub CommandProd_Click()
   If fValidaParte Then
      generarUnParte
      limpiarFormularioParcial
   End If
End Sub
Private Function fValidaParte() As Boolean
   Dim mRet As Boolean
   Dim mMensajeError As String
    
   mRet = True
      
   If Trim(Combo1.Text) = "" Then
      mRet = False
      mMensajeError = "Debe seleccionar un Ramal"
   End If
   
   If mRet Then
      If Trim(Combo2.Text) = "" Then
         mRet = False
         mMensajeError = "Debe seleccionar un Tipo de Activo"
      End If
   End If
      
   If mRet Then
      'Si seleccione un activo de la BD
      If Not Text3.Enabled Then
         If Trim(Combo3.Text) = "" Then
            mRet = False
            mMensajeError = "Debe seleccionar un Acceso"
         End If
         If mRet Then
            If Trim(Combo4.Text) = "" Then
               mRet = False
               mMensajeError = "Debe seleccionar un Km"
            End If
         End If
         If mRet Then
            If Trim(Combo5.Text) = "" Then
               mRet = False
               mMensajeError = "Debe seleccionar una Progresiva"
            End If
         End If
      Else
         If Trim(Text3.Text) = "" Then
            mRet = False
            mMensajeError = "Debe completar el campo 'Otro Activo'"
         End If
      End If
   End If
      
   If mRet Then
      If Trim(Combo6.Text) = "" Then
         mRet = False
         mMensajeError = "Debe seleccionar un 'Problema'."
      End If
   End If
      
      
   If mRet Then
      If Text4.Enabled Then
         If Trim(Text4.Text) = "" Then
            mRet = False
            mMensajeError = "Debe completar el campo  'Otro Problema'."
         End If
      End If
   End If
      
      
   If mRet Then
      If Trim(Combo7.Text) = "" Then
         mRet = False
         mMensajeError = "Debe seleccionar la Prioridad."
      End If
   End If
      
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   fValidaParte = mRet
End Function

Private Sub generarUnParte()
   Dim mCodActivo As String
   Dim mDescActivo As String
   Dim mCodTipoProblema As String
   Dim mDescTipoProblema As String
   Dim mRec1 As New ADODB.Recordset
   
   'Si el campo Otro Activo esta inhabilitado es porque elegi un Activo existe en la BD
   If Not Text3.Enabled Then
      mCodActivo = Trim(Right(Combo5, 20))
      mDescActivo = mObj.sCampoDescrip("COM_Activos", "Codigo = '" & mCodActivo & "'", 1)
   Else
      
      'IMPORTANTE: En la tabla COM_Activos solo puede existir un registro con CodRamal = NULL
      Set mRec1 = mObj.oEjecutarSelect("SELECT * FROM COM_Activos WHERE CodRamal IS NULL ")
      mCodActivo = mRec1!Codigo
      mRec1.Close
      mDescActivo = Trim(Text3.Text)
   End If
      
   mCodTipoProblema = Trim(Right(Combo6, 2))
      
   'Si el campo Otro Activo esta inhabilitado es porque elegi un Tipo de Problema existente en la BD
   If Not Text4.Enabled Then
      mDescTipoProblema = Trim(Left(Combo6, 50))
   Else
      mDescTipoProblema = Trim(Text4.Text)
   End If
        
   FlexPartes.AddItem "X" & vbTab & mDescActivo & vbTab & mDescTipoProblema & vbTab & Trim(Combo7.Text) & vbTab & mCodActivo & vbTab & mCodTipoProblema
   
   Set mRec1 = Nothing

End Sub

Private Sub limpiarFormularioParcial()
   sLlenoTipoActivo
   
   Combo3.Clear
   Combo3.Enabled = False
   
   Combo4.Clear
   Combo4.Enabled = False
   
   Combo5.Clear
   Combo5.Enabled = False
   
   Combo6.Clear
      
   sLlenoPrioridad
      
   Text3.Text = ""
   Text3.Enabled = False
   Text4.Text = ""
   Text4.Enabled = False
   
End Sub
Private Sub limpiarFormularioTotal()
   Dim mi As Integer
   Text1.Text = ""
   sLlenoRamal
   limpiarFormularioParcial
   'Elimino los registros grilla inferior
   For mi = FlexPartes.Rows To 3 Step -1
      FlexPartes.RemoveItem mi
   Next
End Sub

Private Sub FlexPartes_Click()
   If FlexPartes.MouseCol = 0 And FlexPartes.MouseRow > 0 Then
      If FlexPartes.Rows > 2 Then
         FlexPartes.RemoveItem FlexPartes.MouseRow
         FlexPartes.Row = 0
      End If
   End If
End Sub


Private Sub Form_Load()
   Me.Width = 17085
   Me.Height = 9920
   sAlinearForm Me

   
   Combo3.Enabled = False
   Combo4.Enabled = False
   Combo5.Enabled = False
   
   Text3.Enabled = False
   Text4.Enabled = False
   

   sLlenoRamal
   cboRamalListIndex = Combo1.ListIndex

   
   sLlenoTipoActivo
   sLlenoPrioridad
   initPartes
End Sub

Private Sub sLlenoRamal()
   Dim mRec1 As New ADODB.Recordset
   
   Combo1.Clear
   Set mRec1 = mObj.oEjecutarSelect("SELECT Codigo, Descripcion From COM_Ramales ORDER BY Descripcion; ")

   Do While Not mRec1.EOF
      Combo1.AddItem mRec1!descripcion & Space(100) & mRec1!Codigo
      mRec1.MoveNext
   Loop
   
   mRec1.Close
   Set mRec1 = Nothing
End Sub

Private Sub sLlenoPrioridad()
   Combo7.Clear
   Combo7.AddItem "Alta"
   Combo7.AddItem "Media"
   Combo7.AddItem "Baja"
   Combo7.ListIndex = 1
End Sub


Private Sub sLlenoTipoActivo()
   Dim mRec1 As New ADODB.Recordset
   
   Combo2.Clear

   Set mRec1 = mObj.oEjecutarSelect("SELECT Codigo, Descripcion From COM_TiposActivo ORDER BY Descripcion; ")

   Do While Not mRec1.EOF
      Combo2.AddItem mRec1!descripcion & Space(100) & mRec1!Codigo
      mRec1.MoveNext
   Loop
   
   mRec1.Close
   Set mRec1 = Nothing
End Sub

Private Sub initPartes()
   filaAnt = 0
   columnAnt = 0
   Text2.Visible = False
   
   With FlexPartes
      .ColWidth(0) = 200
      .ColWidth(1) = 4720
      .ColWidth(2) = 9470
      .ColWidth(3) = 1650
      .ColWidth(4) = 0
      .ColWidth(5) = 0
    
      .TextMatrix(0, 1) = "Activo"
      .TextMatrix(0, 2) = "Problema"
      .TextMatrix(0, 3) = "Prioridad"
      .TextMatrix(0, 4) = "CodActivo"
      .TextMatrix(0, 5) = "CodProblema"
      
      .ColAlignment(2) = flexAlignLeftCenter

      .RowHeight(1) = 0
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mObj = Nothing
   Set mRec = Nothing
   ShowMenu 47, True, False
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 46 Then
      KeyAscii = fNumeroKeyPress(KeyAscii)
   End If
   
   If KeyAscii = 13 Then
      FlexPartes.TextMatrix(filaAnt, columnAnt) = Text2.Text
      Text2.Visible = False
      FlexPartes.ScrollBars = flexScrollBarVertical
   End If
End Sub

Private Sub Text2_LostFocus()
   If FlexPartes.Col <> columnAnt Or FlexPartes.Row <> filaAnt Then
      'En este caso 3 es la columna que seria editable
      If columnAnt = 3 Then
         FlexPartes.TextMatrix(filaAnt, columnAnt) = Text2.Text
      End If
   End If
End Sub
