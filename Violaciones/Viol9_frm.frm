VERSION 5.00
Begin VB.Form Viol9_frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulo de importación de Violaciones"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7455
   Begin VB.CommandButton Command1 
      Caption         =   "&Volver"
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   8
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Procesar"
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   7
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   3720
         Width           =   6735
      End
      Begin VB.FileListBox File1 
         Height          =   2820
         Left            =   3360
         Pattern         =   "*.xls"
         TabIndex        =   3
         Top             =   480
         Width           =   3615
      End
      Begin VB.DirListBox Dir1 
         Height          =   2565
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   3015
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre de archivo seleccionado"
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
         Left            =   240
         TabIndex        =   6
         Top             =   3525
         Width           =   2805
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Selección de archivo"
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
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Viol9_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObjViol As New clViolaciones
Dim XLS As EXCEL.Application
Dim mRec As New ADODB.Recordset
Dim mI As Integer
Dim aDatos(8) As String
Dim mErr As String
Dim mCodVeh As String
Dim mCodMod As String
Dim mCodCol As String

Private Sub Form_Load()
Dir1.Path = "C:\"
File1.Path = Dir1.Path
sAlinearForm Me
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
Text1.Text = File1.Path & File1.List(File1.ListIndex)
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mOk As Integer
Dim mKO As Integer
If Index = 0 Then
   If Text1.Text <> "" Then
      If MsgBox("¿Está seguro de importar violaciones con el archivo:" & Chr(13) & Text1.Text & "?", vbYesNo, sMessage & " - Atención!!") = vbYes Then
         Set XLS = CreateObject("Excel.Application")
         XLS.Application.WorkBooks.Open filename:=File1.Path & "" & File1.List(File1.ListIndex)
         Open Left(Text1.Text, Len(Text1.Text) - 4) & ".txt" For Output Shared As #1
         For mI = 1 To 8
            aDatos(mI) = ""
         Next
         mI = 2
         Me.MousePointer = 11
         sMsgEspere Me, "Importando violaciones a la base de datos...", True
         mOk = 0
         mKO = 0
         Do While XLS.Cells(mI, 1) <> ""
            aDatos(1) = Trim(XLS.Cells(mI, 1))        'Estacion
            aDatos(2) = UCase(Trim(XLS.Cells(mI, 2))) 'Via
            aDatos(3) = Trim(XLS.Cells(mI, 3))        'Fecha
            aDatos(4) = Trim(XLS.Cells(mI, 4))        'Hora
            aDatos(5) = UCase(Trim(XLS.Cells(mI, 5))) 'Patente
            aDatos(6) = UCase(Trim(XLS.Cells(mI, 6))) 'Vehiculo
            aDatos(7) = UCase(Trim(XLS.Cells(mI, 7))) 'Modelo
            aDatos(8) = UCase(Trim(XLS.Cells(mI, 8))) 'Color
            If CamposOK() Then
               Set mRec = mObjViol.oEjecutarSelect("SELECT * FROM Registros WHERE ESTACION = '" & aDatos(1) & "' And VIA = '" & aDatos(2) & "' And FECHA = '" & mId(aDatos(3), 7, 4) & "-" & mId(aDatos(3), 4, 2) & "-" & mId(aDatos(3), 1, 2) & "' And HORA = '" & aDatos(4) & "' And PATENTE = '" & aDatos(5) & "'")
               If Not mRec.EOF Then
                  mErr = "Ya está registrada dicha violación"
                  mKO = mKO + 1
               Else
                  mOk = mOk + 1
                  mErr = "Importacion OK"
                  mObjViol.xInsRegistros aDatos(3), aDatos(4), aDatos(1), Left(aDatos(2), 2), Right(aDatos(2), 1), aDatos(5), mCodMod, mCodCol, "", "", "", "", "", mCodVeh, "", "V"
               End If
               mRec.Close
            Else
               mKO = mKO + 1
            End If
            Print #1, aDatos(1) & "-" & aDatos(2) & "-" & aDatos(3) & "-" & aDatos(4) & "-" & aDatos(5) & "-" & aDatos(6) & "-" & aDatos(7) & "-" & aDatos(8) & "--->" & mErr
            mI = mI + 1
         Loop
         Close #1
         Set XLS = Nothing
         Me.MousePointer = 0
         sMsgEspere Me, "", False
         MsgBox "Operación Finalizada" & Chr(13) & "Registros Grabados: " & mOk & Chr(13) & "Registros Rechazados: " & mKO & Chr(13) & "Registros Totales: " & mOk + mKO & Chr(13) & "Ver archivo: " & Left(Text1.Text, Len(Text1.Text) - 4) & ".txt", vbInformation, sMessage & " - Atención!!"
      End If
   Else
      MsgBox "Seleccione un archivo para procesar!", vbCritical, "Atención"
   End If
Else
   Unload Me
   ShowMenu 5, True, False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set XLS = Nothing
Set mObjViol = Nothing
Set mRec = Nothing
End Sub

Private Function CamposOK()
Dim mRet As Boolean
mRet = True
mErr = ""
If Len(aDatos(1)) <> 2 Then
   mRet = False
   mErr = "Estacion errónea"
End If
If mRet Then
   If Len(aDatos(2)) <> 3 Then
      mRet = False
      mErr = "Via errónea"
   End If
End If
If mRet Then
   If Len(aDatos(3)) <> 10 Then
      mRet = False
      mErr = "Fecha errónea"
   End If
End If
If mRet Then
   If Len(aDatos(4)) <> 5 Then
      mRet = False
      mErr = "Hora errónea"
   End If
End If
If mRet Then
   If Len(aDatos(5)) <> 6 Then
      mRet = False
      mErr = "Patente errónea"
   End If
End If
If mRet Then
   Set mRec = mObjViol.oEjecutarSelect("SELECT * FROM marcas WHERE DESCRIPCION = '" & aDatos(6) & "' And BAJA IS NULL")
   If mRec.EOF Then
      mRet = False
      mErr = "Marca " & aDatos(6) & " inexistente"
   Else
      mCodVeh = mRec!Codigo
   End If
End If
If mRet Then
   Set mRec = mObjViol.oEjecutarSelect("SELECT * FROM modelos WHERE DESCRIPCION = '" & aDatos(7) & "' And CODMARCA = '" & mCodVeh & "' And BAJA IS NULL")
   If mRec.EOF Then
      mRet = False
      mErr = "Modelo " & aDatos(6) & "-" & aDatos(7) & " inexistente"
   Else
      mCodMod = mRec!Codigo
   End If
   mRec.Close
End If
If mRet Then
   Set mRec = mObjViol.oEjecutarSelect("SELECT * FROM colores WHERE DESCRIPCION = '" & aDatos(8) & "' And BAJA IS NULL")
   If mRec.EOF Then
      mRet = False
      mErr = "Color " & aDatos(8) & " inexistente"
   Else
      mCodCol = mRec!Codigo
   End If
   mRec.Close
End If
CamposOK = mRet
End Function
