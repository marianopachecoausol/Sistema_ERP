VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   Caption         =   "Form1"
   ClientHeight    =   9780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   ScaleHeight     =   9780
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   5520
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   3855
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3480
      Width           =   9615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SFTP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002170E7&
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim oShell As WshShell
Dim oExec As WshExec
Dim ret As String
Dim mComando As String
Dim ejecutar_Dos
    mComando = App.Path & "\pscp.exe -pw AdminGCO2010$ -sftp c:\auoesteAR.xml ausolwaze@52.162.163.73:/home/ausolwaze"
'    Replace mComando, "(", "\("
'    Replace mComando, ")", "\)"
'    mComando = "'" & mComando & "'"
    
    MsgBox mComando
    Set oShell = New WshShell
    DoEvents

    ' ejecutar el comando
    Set oExec = oShell.Exec("%comspec% /c " & mComando)
    ret = oExec.StdOut.ReadAll()

    ' retornar la salida y devolverla a la función
    Text1.Text = ret  ' Replace(ret, Chr(10), vbNewLine)

    DoEvents
    Me.SetFocus

End Sub

'Private Sub Command2_Click()
'Dim mObj As New clRNov
'Dim mRec As New ADODB.Recordset
'
'
'   Set mRec = mObj.oCallStore("getCoordenadasWaze")
'   If Not mRec.EOF Then
'
'   End If
'
'   Set mObj = Nothing
'   Set mRec = Nothing
'End Sub

