VERSION 5.00
Begin VB.Form Viol7_frm 
   BackColor       =   &H00CECECE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulo de Actualización de Direcciones y Envíos de Cartas Documentos."
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9120
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   9120
   Begin VB.CommandButton Command1 
      Caption         =   "Actualizar Envíos de Cartas Documentos"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   6480
      TabIndex        =   8
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualizar Direcciones"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   6480
      TabIndex        =   7
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CECECE&
      Height          =   5055
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   4620
         Width           =   5895
      End
      Begin VB.FileListBox File1 
         Height          =   3795
         Left            =   3120
         Pattern         =   "*.xls"
         TabIndex        =   4
         Top             =   480
         Width           =   3015
      End
      Begin VB.DirListBox Dir1 
         Height          =   3465
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   2775
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00CECECE&
         Caption         =   "Nombre:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   4400
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00CECECE&
         Caption         =   "Selección de Archivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1920
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C1DBD8&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Seleccionar el archivo de Excel para la actualización elegida y luego presionar sobre el botón correspondiente."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6600
      TabIndex        =   9
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "Viol7_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObjViol As New clViolaciones
Dim XLS As EXCEL.Application
Dim mRec As New ADODB.Recordset
Dim mI As Integer

Private Sub Form_Load()
Dir1.Path = "C:\"
File1.Path = Dir1.Path
Me.Height = 5565
Me.Width = 9240
sAlinearForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set XLS = Nothing
Set mObjViol = Nothing
Set mRec = Nothing
ShowMenu 5, True, False
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mJ As Integer
Dim mDatos(6) As String
Dim mFlag As Boolean
If MsgBox("¿Está seguro que desea Actualizar con el archivo:" & Chr(13) & File1.Path & "\" & File1.List(File1.ListIndex) & "?", vbYesNo, sMessage & " - Atención!!") = vbYes Then
   Set XLS = CreateObject("Excel.Application")
   XLS.Application.WorkBooks.Open filename:=File1.Path & "" & File1.List(File1.ListIndex)
   For mI = 1 To 6
      mDatos(mI) = ""
   Next
   mI = 2
   Me.MousePointer = 11
   sMsgEspere Me, "Pasando información a la base de datos.", True
   Select Case Index
      Case 0 'Direcciones
         Do While XLS.Cells(mI, 1) <> ""
            If XLS.Cells(mI, 2) <> "" And XLS.Cells(mI, 3) <> "" Then
               For mJ = 1 To 5
                  mDatos(mJ) = Trim(XLS.Cells(mI, mJ))
               Next
               mI = mI + 1
               Set mRec = mObjViol.oTabla("provincias", "where descripcion like '%" & mDatos(4) & "%'")
               If Not mRec.EOF Then
                  mDatos(4) = mRec!Codigo
               Else
                  mDatos(4) = "01"
               End If
               mRec.Close
               Set mRec = mObjViol.oTabla("direcciones", "where patente = '" & Trim(mDatos(1)) & "'")
               If Not mRec.EOF Then
                  mFlag = mObjViol.xUpDirecciones(mDatos(2), mDatos(3), mDatos(4), Left(mDatos(5), 4), mDatos(1))
               Else
                  mFlag = mObjViol.xInsDirecciones(mDatos(1), mDatos(2), mDatos(3), mDatos(4), mDatos(5))
               End If
               mRec.Close
            Else
               mI = mI + 1
            End If
         Loop
         Set XLS = Nothing
      Case 1 'Envios
         Do While XLS.Cells(mI, 6) <> ""
            If XLS.Cells(mI, 7) <> "" And XLS.Cells(mI, 9) <> "" Then
               mDatos(1) = Trim(XLS.Cells(mI, 6)) 'patente
               mDatos(2) = Trim(XLS.Cells(mI, 9)) 'nrocarta
               mDatos(3) = Trim(XLS.Cells(mI, 7)) 'fecha
               mDatos(4) = Trim(XLS.Cells(mI, 8)) 'codentrega
               mDatos(6) = Trim(XLS.Cells(mI, 10)) 'Tipo (V / D)
               mI = mI + 1
               Set mRec = mObjViol.oTabla("entregas", "where descripcion like '%" & mDatos(4) & "%'")
               If Not mRec.EOF Then
                  mDatos(4) = mRec!Codigo
               Else
                  mDatos(4) = "99"
               End If
               mRec.Close
               Set mRec = mObjViol.oEnviosxPatente(Trim(mDatos(1)) & "' and fecha='" & Format(Trim(mDatos(3)), "yyyy-mm-dd"), mDatos(6))
               If Not mRec.EOF Then
                  If mRec!Fecha <> Null Or mRec!Fecha <> "" Then
                     mFlag = mObjViol.xUpEnvios(mDatos(2), mDatos(3), mDatos(4), mDatos(1), mDatos(3), mDatos(6))
                  End If
               Else
                  mFlag = mObjViol.xInsEnvios(mDatos(1), mDatos(2), Trim(mDatos(3)), mDatos(4), "", mDatos(6))
               End If
               mRec.Close
            Else
               mI = mI + 1
            End If
         Loop
         MsgBox "Operación Finalizada", vbInformation, sMessage & " - Atención!!"
         Set XLS = Nothing
   End Select
   Me.MousePointer = 0
   sMsgEspere Me, "", False
   MsgBox "Operación Finalizada", vbInformation, sMessage
End If
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
On Error Resume Next
File1.Path = Dir1.Path
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()
Text1.Text = File1.Path & "\" & File1.List(File1.ListIndex)
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Then
   Command1(0).Enabled = True
   Command1(1).Enabled = True
Else
   Command1(0).Enabled = False
   Command1(1).Enabled = False
End If
End Sub
