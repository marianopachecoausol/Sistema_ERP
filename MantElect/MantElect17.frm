VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MantElect17 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelacion de tareas (partes)"
   ClientHeight    =   9405
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   16965
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   16965
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Index           =   1
      Left            =   9450
      TabIndex        =   4
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Confirmar"
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   3
      Top             =   8880
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8730
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   120
      Width           =   16920
      Begin VB.Frame Frame2 
         Caption         =   "Motivo de la Cancelación"
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
         Height          =   1215
         Left            =   360
         TabIndex        =   13
         Top             =   120
         Width           =   16275
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   1560
            MaxLength       =   90
            TabIndex        =   14
            Top             =   500
            Width           =   9735
         End
         Begin VB.Label Label1 
            Caption         =   "Motivo:"
            Height          =   255
            Left            =   720
            TabIndex        =   15
            Top             =   585
            Width           =   735
         End
      End
      Begin VB.CommandButton CommandPartes 
         Height          =   375
         Index           =   1
         Left            =   8280
         Picture         =   "MantElect17.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5400
         Width           =   330
      End
      Begin VB.CommandButton CommandPartes 
         Height          =   375
         Index           =   0
         Left            =   7642
         Picture         =   "MantElect17.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5400
         Width           =   330
      End
      Begin VB.Frame Frame14 
         Caption         =   "Búsqueda de Partes:"
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
         Height          =   3855
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   16275
         Begin VB.ComboBox Combo7 
            Height          =   315
            Left            =   6840
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   360
            Width           =   2415
         End
         Begin VB.ComboBox Combo6 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   360
            Width           =   3255
         End
         Begin MSFlexGridLib.MSFlexGrid FlexPartes 
            Height          =   2805
            Left            =   240
            TabIndex        =   10
            Top             =   960
            Width           =   15735
            _ExtentX        =   27755
            _ExtentY        =   4948
            _Version        =   327680
            Cols            =   9
         End
         Begin VB.Label Label13 
            Caption         =   "Detalle:"
            Height          =   255
            Left            =   6000
            TabIndex        =   8
            Top             =   465
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "Origen:"
            Height          =   255
            Left            =   720
            TabIndex        =   6
            Top             =   465
            Width           =   975
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Partes para cancelar"
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
         Height          =   2895
         Left            =   360
         TabIndex        =   1
         Top             =   5760
         Width           =   16275
         Begin MSFlexGridLib.MSFlexGrid FlexPartACancelar 
            Height          =   2370
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   15735
            _ExtentX        =   27755
            _ExtentY        =   4180
            _Version        =   327680
            Cols            =   9
         End
      End
   End
End
Attribute VB_Name = "MantElect17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mObj As New clMantElect
Dim mObjInven As New clInven
Dim mRec As New ADODB.Recordset
Dim mRenglonPartes As Integer
Dim mRenglonPartAsignados As Integer

Dim filaAnt As Integer
Dim columnAnt As Integer


'TODO: Ver si es necesario utilizar las siguientes variables:
Dim mCodParte As Integer

Dim cboListIndex As Integer

Dim mEsOTcerrada As Boolean

Dim cboOrigenListIndex As Integer
Dim cboDetalleListIndex As Integer

Dim vPartesOriginal() As Double





Private Sub Combo7_Click()

   Dim mi As Integer
   Dim mj As Integer
   Dim mNroComunicado As String
   Dim mTramo As String
   Dim mRamal As String
   Dim Origen As String
   Dim sListaPartesSeleccionados As String
   
   sListaPartesSeleccionados = "-1"
   
   eliminoGrillaPartes
   
   'If cboDetalleListIndex <> Combo7.ListIndex And FlexPartACancelar.Rows > 2 Then
   
   If FlexPartACancelar.Rows > 2 Then
      For mj = 2 To FlexPartACancelar.Rows - 1
         sListaPartesSeleccionados = sListaPartesSeleccionados & "," & FlexPartACancelar.TextMatrix(mj, 1)
      Next
   End If
   
   
   
   If cboDetalleListIndex <> Combo7.ListIndex Then
            eliminoGrillaPartes
            'eliminoGrillaPartesAsignados
            Origen = Trim(Right(Combo6.Text, 4))
            Select Case Origen
               Case "OPE"
                  mTramo = Trim(Left(Combo7.Text, 2))
                  cargarGrillaConPartesOperaciones mTramo, sListaPartesSeleccionados
               Case "REL"
                  mRamal = Trim(Left(Combo7.Text, 50))
                  cargarGrillaConPartesDeRelevamientos mRamal, sListaPartesSeleccionados
               Case "COM"
                  mNroComunicado = Trim(Combo7.Text)
                  cargarGrillaConPartesDeComunicado mNroComunicado, sListaPartesSeleccionados
            End Select
         
         cboDetalleListIndex = Combo7.ListIndex
   End If
   
   cboDetalleListIndex = -99
   
End Sub

Private Sub preparaArrayPartes(ByRef pvParte_NroParte() As Double, ByRef pvParte_CodEdificio() As String, ByRef pvParte_Descripcion() As String, ByRef pvParte_SecAire() As Integer, ByRef pvParte_OpGen() As String, ByRef pvParte_CodSuperv() As String)
   Dim mj As Integer
   Dim cantPartes As Integer

   cantPartes = FlexPartACancelar.Rows - 2
   If cantPartes > 0 Then
      ReDim pvParte_NroParte(0 To cantPartes - 1) As Double
      ReDim pvParte_CodEdificio(0 To cantPartes - 1) As String
      ReDim pvParte_Descripcion(0 To cantPartes - 1) As String
      ReDim pvParte_SecAire(0 To cantPartes - 1) As Integer
      ReDim pvParte_OpGen(0 To cantPartes - 1) As String
      ReDim pvParte_CodSuperv(0 To cantPartes - 1) As String
      
           
      For mj = 2 To FlexPartACancelar.Rows - 1
        pvParte_NroParte(mj - 2) = FlexPartACancelar.TextMatrix(mj, 1)
        pvParte_CodEdificio(mj - 2) = FlexPartACancelar.TextMatrix(mj, 3)
        pvParte_Descripcion(mj - 2) = FlexPartACancelar.TextMatrix(mj, 4)
        pvParte_SecAire(mj - 2) = FlexPartACancelar.TextMatrix(mj, 6)
        pvParte_OpGen(mj - 2) = FlexPartACancelar.TextMatrix(mj, 7)
        pvParte_CodSuperv(mj - 2) = FlexPartACancelar.TextMatrix(mj, 8)
      Next
   Else
      'Esta linea sirve como flag para avisar en el procedimiento xinsOT que no tiene registros y por ende no me de error de intervalos al usar Ubound y Lubound
      'Igualmente no deberia pasar debido a que no antes de confirmar valido si hay algun parte elegido.
      ReDim pvParte_NroParte(0)
      pvParte_NroParte(0) = 0
   End If
End Sub




Private Sub preparaArrayPartesold(ByRef pvPartes_Cancel() As Integer)
   Dim mj As Integer
   Dim cantPartes As Integer

   cantPartes = FlexPartACancelar.Rows - 2
   If cantPartes > 0 Then
      ReDim pvPartes_Cancel(0 To cantPartes - 1) As Integer
         
      For mj = 2 To FlexPartACancelar.Rows - 1
         pvPartes_Cancel(mj - 2) = FlexPartACancelar.TextMatrix(mj, 1)
      Next
   Else
      'Esta linea sirve como flag para avisar en el procedimiento xinsOT que no tiene registros y por ende no me de error de intervalos al usar Ubound y Lubound
      'Igualmente no deberia pasar debido a que no antes de confirmar valido si hay algun parte elegido.
      ReDim pvPartes_Cancel(0)
      pvPartes_Cancel(0) = 0
   End If
End Sub

Private Sub Command2_Click(Index As Integer)
   Dim vPartes_Cancel() As Double
   Dim vParte_CodEdificio() As String
   Dim vParte_Descripcion() As String
   Dim vParte_SecAire() As Integer
   Dim vParte_OpGen() As String
   Dim vParte_CodSuperv() As String
   If Index = 0 Then
      If fValidaCancelacion Then
         preparaArrayPartes vPartes_Cancel(), vParte_CodEdificio(), vParte_Descripcion(), vParte_SecAire(), vParte_OpGen(), vParte_CodSuperv()
         mObj.xCancelarPartes Trim(Text1.Text), vPartes_Cancel(), vParte_CodEdificio(), vParte_Descripcion(), vParte_SecAire(), vParte_OpGen(), vParte_CodSuperv(), Trim(Right(MDI.mUser, 15))
         limpiarFormulario
      End If
   Else
      Unload Me
   End If
End Sub

Private Sub limpiarFormulario()
   Dim mi As Integer
   Text1.Text = ""
   'Elimino los registros grilla inferior
   For mi = FlexPartACancelar.Rows To 3 Step -1
      FlexPartACancelar.RemoveItem mi
   Next
End Sub


Private Function fValidaCancelacion() As Boolean
   Dim mRet As Boolean
   Dim mMensajeError As String

   mRet = True
   
   If Trim(Text1.Text) = "" Then
      mRet = False
      mMensajeError = "Debe completar el Motivo"
   End If

   If mRet Then
      If FlexPartACancelar.Rows <= 2 Then
         mRet = False
         mMensajeError = "Al menos se debe seleccionar un Parte"
      End If
   End If
   
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If

   fValidaCancelacion = mRet
End Function







Private Sub CommandPartes_Click(Index As Integer)

   Dim sListaPartesSeleccionados As String
   Dim mj As Integer
   Dim Origen As String
   Dim mTramo As String
   Dim mRamal As String
   Dim mNroComunicado As String
      
   sListaPartesSeleccionados = "-1"
'   If FlexPartACancelar.Rows > 2 Then
'      For mj = 2 To FlexPartACancelar.Rows - 1
'         sListaPartesSeleccionados = sListaPartesSeleccionados & "," & FlexPartACancelar.TextMatrix(mj, 1)
'      Next
'   End If
'
   If Index = 0 Then
      If mRenglonPartes > 0 Then
         If Trim(FlexPartes.TextMatrix(mRenglonPartes, 1)) <> "" Then
            
            FlexPartACancelar.AddItem vbTab & FlexPartes.TextMatrix(mRenglonPartes, 1) & vbTab & FlexPartes.TextMatrix(mRenglonPartes, 2) & vbTab & FlexPartes.TextMatrix(mRenglonPartes, 3) & vbTab & FlexPartes.TextMatrix(mRenglonPartes, 4) & vbTab & FlexPartes.TextMatrix(mRenglonPartes, 5) & vbTab & FlexPartes.TextMatrix(mRenglonPartes, 6) & vbTab & FlexPartes.TextMatrix(mRenglonPartes, 7) & vbTab & FlexPartes.TextMatrix(mRenglonPartes, 8)
            'MsgBox FlexPartACancelar.Rows
         End If
         
         If FlexPartes.Rows > 2 Then
            FlexPartes.RemoveItem mRenglonPartes
         
            mRenglonPartes = 0
         Else
            If Trim(FlexPartes.TextMatrix(mRenglonPartes, 1)) <> "" Then
               FlexPartes.TextMatrix(mRenglonPartes, 1) = ""
               FlexPartes.TextMatrix(mRenglonPartes, 2) = ""
         
               mRenglonPartes = 0
            End If
         End If
      End If
   Else
      If FlexPartACancelar.Rows > 2 And mRenglonPartAsignados > 1 Then
         FlexPartACancelar.RemoveItem (mRenglonPartAsignados)
         
         If FlexPartACancelar.Rows > 2 Then
            For mj = 2 To FlexPartACancelar.Rows - 1
               sListaPartesSeleccionados = sListaPartesSeleccionados & "," & FlexPartACancelar.TextMatrix(mj, 1)
            Next
         End If
            
         mRenglonPartes = 0
'
'         FlexPartes.Clear
'         'Elimino los registros  de la grilla superior
'         For mj = FlexPartes.Rows To 3 Step -1
'            FlexPartes.RemoveItem mj
'         Next

         eliminoGrillaPartes
         
         
         
         With FlexPartes
            .TextMatrix(0, 1) = "Parte"
            .TextMatrix(0, 2) = "Fecha Solicitud"
            .TextMatrix(0, 3) = "Lugar"
            .TextMatrix(0, 4) = "Descripcion de la Solicitud"
            .TextMatrix(0, 5) = "Prioridad"
            
            .TextMatrix(0, 6) = "Sector Aire"
            
            .RowHeight(1) = 0
         End With
        
         If Combo6.ListIndex >= 0 Then
            
            Origen = Trim(Right(Combo6.Text, 3))
            Select Case Origen
               Case "OPE"
                  mTramo = Trim(Left(Combo7.Text, 2))
                  cargarGrillaConPartesOperaciones mTramo, sListaPartesSeleccionados
               Case "REL"
                  mRamal = Trim(Left(Combo7.Text, 50))
                  cargarGrillaConPartesDeRelevamientos mRamal, sListaPartesSeleccionados
               Case "COM"
                  mNroComunicado = Trim(Combo7.Text)
                  cargarGrillaConPartesDeComunicado mNroComunicado, sListaPartesSeleccionados
            End Select
         End If
         

      End If
      mRenglonPartAsignados = 0
   End If



End Sub


Private Sub FlexPartACancelar_Click()
   Dim mi As Integer
   
   If FlexPartACancelar.MouseRow > 0 Then
   
      If mRenglonPartAsignados <> 0 Then
         FlexPartACancelar.Row = mRenglonPartAsignados
         For mi = 1 To FlexPartACancelar.Cols - 1
            FlexPartACancelar.Col = mi
            FlexPartACancelar.CellBackColor = vbWhite
         Next
      End If
      
      mRenglonPartAsignados = FlexPartACancelar.MouseRow
   
      FlexPartACancelar.Row = mRenglonPartAsignados
      For mi = 1 To FlexPartACancelar.Cols - 1
         FlexPartACancelar.Col = mi
         FlexPartACancelar.CellBackColor = &H80000003
      Next
      
      If mRenglonPartAsignados > 1 Then
          mCodParte = FlexPartACancelar.TextMatrix(mRenglonPartAsignados, 1)
      End If
   Else
      FlexPartACancelar.Row = mRenglonPartAsignados
      For mi = 1 To FlexPartACancelar.Cols - 1
         FlexPartACancelar.Col = mi
         FlexPartACancelar.CellBackColor = vbWhite
      Next
      mRenglonPartAsignados = 0
   End If

End Sub

Private Sub FlexPartes_Click()

   Dim mi As Integer
   
   If FlexPartes.MouseRow > 0 Then
   
      If mRenglonPartes <> 0 Then
         FlexPartes.Row = mRenglonPartes
         For mi = 1 To FlexPartes.Cols - 1
            FlexPartes.Col = mi
            FlexPartes.CellBackColor = vbWhite
         Next
      End If
      
      mRenglonPartes = FlexPartes.MouseRow
   
      FlexPartes.Row = mRenglonPartes
      For mi = 1 To FlexPartes.Cols - 1
         FlexPartes.Col = mi
         FlexPartes.CellBackColor = &H80000003
      Next
      
      If mRenglonPartes > 1 Then
          mCodParte = FlexPartes.TextMatrix(mRenglonPartes, 1)
      End If
   Else
      FlexPartes.Row = mRenglonPartes
      For mi = 1 To FlexPartes.Cols - 1
         FlexPartes.Col = mi
         FlexPartes.CellBackColor = vbWhite
      Next
      mRenglonPartes = 0
   End If

End Sub


Private Sub Text10_KeyPress(KeyAscii As Integer)
      If KeyAscii <> 46 Then
         KeyAscii = fNumeroKeyPress(KeyAscii)
      End If
End Sub

Private Sub text8_KeyPress(KeyAscii As Integer)
         KeyAscii = fNumeroKeyPress(KeyAscii)
End Sub


Private Sub Form_Load()

   eliminoGrillaPartes
   'InicializoCboOrigen
   sLlenoCboOrigen
   InicializoCboDetalle
   
   Me.Width = 17090
   Me.Height = 9920
   sAlinearForm Me
   
   'sLlenoCboOrigen
   cboOrigenListIndex = -99
   cboDetalleListIndex = -99
   'InicializoCboOrigen

   initPartes

End Sub

Private Sub initPartes()
   mRenglonPartes = 0
   mRenglonPartAsignados = 0
   
   With FlexPartes
      .ColWidth(0) = 200
      .ColWidth(1) = 500
      .ColWidth(2) = 2000
      .ColWidth(3) = 3000
      .ColWidth(4) = 8800
      .ColWidth(5) = 750
      .ColWidth(6) = 0
      .ColWidth(7) = 0
      .ColWidth(8) = 0

      .TextMatrix(0, 1) = "Parte"
      .TextMatrix(0, 2) = "Fecha Solicitud"
      .TextMatrix(0, 3) = "Lugar"
      .TextMatrix(0, 4) = "Descripcion de la Solicitud"
      .TextMatrix(0, 5) = "Prioridad"
      .TextMatrix(0, 6) = "Sector Aire"
      .TextMatrix(0, 7) = "Op.Gen"
      .TextMatrix(0, 8) = "CodSuperv"
      
      .ColAlignment(4) = flexAlignLeftCenter
      
      .RowHeight(1) = 0
   End With
   
   With FlexPartACancelar
      .ColWidth(0) = 200
      .ColWidth(1) = 500
      .ColWidth(2) = 2000
      .ColWidth(3) = 3000
      .ColWidth(4) = 8800
      .ColWidth(5) = 750
      .ColWidth(6) = 0
      .ColWidth(7) = 0
      .ColWidth(8) = 0
      
      .TextMatrix(0, 1) = "Parte"
      .TextMatrix(0, 2) = "Fecha Solicitud"
      .TextMatrix(0, 3) = "Lugar"
      .TextMatrix(0, 4) = "Descripcion de la Solicitud"
      .TextMatrix(0, 5) = "Prioridad"
      .TextMatrix(0, 6) = "Sector Aire"
      .TextMatrix(0, 7) = "Op.Gen"
      .TextMatrix(0, 8) = "CodSuperv"
      
      .ColAlignment(4) = flexAlignLeftCenter
      
      .RowHeight(1) = 0
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mObj = Nothing
   Set mRec = Nothing
   ShowMenu 47, True, False
End Sub


Private Sub sLlenoCboOrigen()
   Combo6.Clear
   
   Combo6.AddItem "OPERACIONES" & Space(50) & "OPE"
   If mObj.esSupervisorElectrico(Trim(Right(MDI.mUser, 20))) Then
      Combo6.AddItem "RELEVAMIENTOS" & Space(50) & "REL"
      Combo6.AddItem "COMUNICADOS" & Space(50) & "COM"
   End If
   Combo6.ListIndex = -1
End Sub

Private Sub sLlenoCboDetalle()
   Dim mRec1 As New ADODB.Recordset
   Dim Origen As String
   
   Combo7.Enabled = True
   Combo7.Clear
   
   Origen = Trim(Right(Combo6.Text, 3))
   
   Select Case Origen
      Case "OPE"
         Set mRec1 = mObj.oEjecutarSelect("SELECT DISTINCT Tramo FROM MantElect.Edificios order by Tramo; ")
         Do While Not mRec1.EOF
            Combo7.AddItem mRec1!Tramo
            mRec1.MoveNext
         Loop
         mRec1.Close
      
      Case "REL"
         Set mRec1 = mObj.oEjecutarSelect("SELECT Codigo, Descripcion FROM COM_Ramales order by Descripcion; ")
         Do While Not mRec1.EOF
            Combo7.AddItem mRec1!descripcion & Space(50) & mRec1!Codigo
            mRec1.MoveNext
         Loop
         mRec1.Close
      Case "COM"
         Set mRec1 = mObj.oEjecutarSelect("SELECT NroComunicado FROM MantElect.COM_Comunicados_H order by Fecha Desc; ")
         Do While Not mRec1.EOF
            Combo7.AddItem mRec1!NroComunicado
            mRec1.MoveNext
         Loop
         mRec1.Close
   End Select
   
   Combo7.ListIndex = -1
End Sub


Private Sub InicializoCboOrigen()
   Combo6.Clear
   Combo6.Enabled = False
   cboOrigenListIndex = -99
End Sub

Private Sub InicializoCboDetalle()
   Combo7.Clear
   Combo7.Enabled = False
End Sub

Private Sub Combo6_Click()

   'If cboOrigenListIndex <> Combo6.ListIndex And FlexPartACancelar.Rows > 2 Then
   If cboOrigenListIndex <> Combo6.ListIndex Then
            eliminoGrillaPartes
            'eliminoGrillaPartesAsignados
            sLlenoCboDetalle
            cboOrigenListIndex = Combo6.ListIndex
            cboDetalleListIndex = -99
   End If

'   If cboOrigenListIndex <> Combo6.ListIndex And FlexPartACancelar.Rows > 2 Then
'         If MsgBox("Si selecciona otro Origen se perderán los partes cargados hasta el momento en la grilla inferior. ¿ Desea continuar ? ", vbYesNo, "Origen") = vbYes Then
'            eliminoGrillaPartes
'            eliminoGrillaPartesAsignados
'            sLlenoCboDetalle
'         Else
'            Combo3.ListIndex = cboOrigenListIndex
'         End If
'         cboOrigenListIndex = Combo3.ListIndex
'   Else
'      If Combo3.ListIndex <> cboOrigenListIndex Then
'         Combo4.Enabled = True
'         eliminoGrillaPartes
'         eliminoGrillaPartesAsignados
'         sLlenoCboDetalle
'      End If
'      cboOrigenListIndex = Combo3.ListIndex
'   End If


End Sub


Private Sub eliminoGrillaPartes()
   Dim mi As Integer
   'Elimino los registros grilla superior
   For mi = FlexPartes.Rows To 3 Step -1
      FlexPartes.RemoveItem mi
   Next
   mRenglonPartes = 0
End Sub




Private Sub cargarGrillaConPartesOperaciones(ByVal pTramo As String, ByVal plistaPartesSeleccionados As String)

   'IMPORTANTE: El parametro plistaPartesSeleccionados no puede venir vacio porque da error, en tal caso se lo fuerza con el parte = -1
   Dim mj As Integer


   
   

                                
                                
'''Backup sentencia igual a la siguiente (por las dudas)
'''   Set mRec = mObj.oEjecutarSelect(" SELECT * FROM ( " & _
'''                                          " SELECT DISTINCT R.Parte,FechaSolic,CodEdificio,descripcion,Prioridad,R.SectorAire,R.FechaIniAsist, concat(U.nombres, ' ',U.apellido ) as OpGen, R.CodSuperv  " & _
'''                                          " FROM Registros R " & _
'''                                             " Inner Join MailsxElectrico M ON R.SectorAire = M.SectorAire " & _
'''                                             " Left Join  OT_Partes OT ON OT.Parte = R.Parte " & _
'''                                             " Left Join  COM_Comunicados_Det C ON C.Parte = R.Parte " & _
'''                                             " Left Join Cancelaciones_Det CNL ON CNL.Parte = R.Parte " & _
'''                                             " Left Join loguser.usuarios U ON U.codusuario = R.OpGen " & _
'''                                          " WHERE Estado NOT IN ('A', 'T') AND Origen = 'O' AND M.codusuario = '" & Trim(Right(MDI.mUser, 20)) & "' " & _
'''                                          " AND (OT.Parte IS NULL OR (OT.Cancelado = 1 AND OT.Finalizado = 'NO' AND NOT EXISTS(SELECT 1 FROM OT_Partes WHERE Parte = OT.Parte AND Cancelado = 0))) " & _
'''                                          " AND C.Parte IS NULL " & _
'''                                          " AND CNL.Parte IS NULL " & _
'''                                          " AND R.CodEdificio like '" & pTramo & "%' " & _
'''                                  " ) AUX WHERE AUX.Parte NOT IN (" & plistaPartesSeleccionados & ") " & _
'''                                  "ORDER BY Parte;")
                                
   Set mRec = mObj.oEjecutarSelect(" SELECT * FROM ( " & _
                                          " SELECT DISTINCT R.Parte,FechaSolic,CodEdificio,descripcion,Prioridad,R.SectorAire,R.FechaIniAsist, concat(U.nombres, ' ',U.apellido ) as OpGen, R.CodSuperv " & _
                                          " FROM Registros R " & _
                                             " Inner Join MailsxElectrico M ON R.SectorAire = M.SectorAire " & _
                                             " Left Join  OT_Partes OT ON OT.Parte = R.Parte " & _
                                             " Left Join  COM_Comunicados_Det C ON C.Parte = R.Parte " & _
                                             " Left Join Cancelaciones_Det CNL ON CNL.Parte = R.Parte " & _
                                             " Left Join loguser.usuarios U ON U.codusuario = R.OpGen " & _
                                          " WHERE Estado NOT IN ('A', 'T') AND Origen = 'O' AND M.codusuario = '" & Trim(Right(MDI.mUser, 20)) & "' " & _
                                          " AND (OT.Parte IS NULL OR (OT.Cancelado = 1 AND OT.Finalizado = 'NO' AND NOT EXISTS(SELECT 1 FROM OT_Partes WHERE Parte = OT.Parte AND Cancelado = 0))) " & _
                                          " AND C.Parte IS NULL " & _
                                          " AND CNL.Parte IS NULL " & _
                                          " AND R.CodEdificio like '" & pTramo & "%' " & _
                                  " ) AUX WHERE AUX.Parte NOT IN (" & plistaPartesSeleccionados & ") " & _
                                  "ORDER BY Parte;")
'
         If Not mRec.EOF Then
            mj = 1
            Do While Not mRec.EOF
               mj = mj + 1
               With FlexPartes
                  .AddItem ""
                  .TextMatrix(mj, 1) = mRec!Parte
                  .TextMatrix(mj, 2) = NVL(mRec!FechaSolic, "")
                  .TextMatrix(mj, 3) = NVL(mRec!CodEdificio, "")
                  .TextMatrix(mj, 4) = NVL(mRec!descripcion, "")
                  .TextMatrix(mj, 5) = NVL(mRec!Prioridad, "")
                  .TextMatrix(mj, 6) = mRec!SectorAire
                  .TextMatrix(mj, 7) = mRec!OpGen
               End With
               mRec.MoveNext
            Loop
         End If
         mRec.Close
End Sub

Private Sub cargarGrillaConPartesDeComunicado(ByVal pNroComunicado As String, ByVal plistaPartesSeleccionados As String)
   'IMPORTANTE: El parametro plistaPartesSeleccionados no puede venir vacio porque da error, en tal caso se lo fuerza con el parte = -1
   Dim mj As Integer


'''Backup sentencia igual a la siguiente (por las dudas)
'''   Set mRec = mObj.oEjecutarSelect(" SELECT * FROM ( " & _
'''                                          " SELECT DISTINCT CD.NroComunicado,R.Parte,FechaSolic,CodEdificio,descripcion,Prioridad,R.SectorAire, FechaIniAsist, concat(U.nombres, ' ',U.apellido ) as OpGen, R.CodSuperv  " & _
'''                                          " FROM Registros R " & _
'''                                             " Inner Join MailsxElectrico M ON R.SectorAire = M.SectorAire " & _
'''                                             " Left Join  OT_Partes OT ON OT.Parte = R.Parte " & _
'''                                             " Inner Join COM_Comunicados_Det CD ON CD.Parte = R.Parte " & _
'''                                             " Inner Join COM_Comunicados_H CH ON CD.NroComunicado = CH.NroComunicado " & _
'''                                             " Left Join Cancelaciones_Det CNL ON CNL.Parte = R.Parte" & _
'''                                             " Left Join loguser.usuarios U ON U.codusuario = R.OpGen " & _
'''                                          " WHERE Estado NOT IN ('A', 'T') " & _
'''                                          " AND Origen = 'O' AND M.codusuario = '" & Trim(Right(MDI.mUser, 20)) & "' " & _
'''                                          " AND (OT.Parte IS NULL OR (OT.Cancelado = 1 " & _
'''                                          " AND OT.Finalizado = 'NO' AND NOT EXISTS(SELECT 1 FROM OT_Partes WHERE Parte = OT.Parte AND Cancelado = 0))) " & _
'''                                          " AND CH.NroComunicado = '" & pNroComunicado & "'" & _
'''                                          " AND CNL.Parte IS NULL " & _
'''                                  " ) AUX WHERE AUX.Parte NOT IN (" & plistaPartesSeleccionados & ") " & _
'''                                  "ORDER BY Parte;")


   Set mRec = mObj.oEjecutarSelect(" SELECT * FROM ( " & _
                                          " SELECT DISTINCT CD.NroComunicado,R.Parte,FechaSolic,CodEdificio,descripcion,Prioridad,R.SectorAire, FechaIniAsist, concat(U.nombres, ' ',U.apellido ) as OpGen, R.CodSuperv " & _
                                          " FROM Registros R " & _
                                             " Inner Join MailsxElectrico M ON R.SectorAire = M.SectorAire " & _
                                             " Left Join  OT_Partes OT ON OT.Parte = R.Parte " & _
                                             " Inner Join COM_Comunicados_Det CD ON CD.Parte = R.Parte " & _
                                             " Inner Join COM_Comunicados_H CH ON CD.NroComunicado = CH.NroComunicado " & _
                                             " Left Join Cancelaciones_Det CNL ON CNL.Parte = R.Parte" & _
                                             " Left Join loguser.usuarios U ON U.codusuario = R.OpGen " & _
                                          " WHERE Estado NOT IN ('A', 'T') " & _
                                          " AND Origen = 'O' AND M.codusuario = '" & Trim(Right(MDI.mUser, 20)) & "' " & _
                                          " AND (OT.Parte IS NULL OR (OT.Cancelado = 1 " & _
                                          " AND OT.Finalizado = 'NO' AND NOT EXISTS(SELECT 1 FROM OT_Partes WHERE Parte = OT.Parte AND Cancelado = 0))) " & _
                                          " AND CH.NroComunicado = '" & pNroComunicado & "'" & _
                                          " AND CNL.Parte IS NULL " & _
                                  " ) AUX WHERE AUX.Parte NOT IN (" & plistaPartesSeleccionados & ") " & _
                                  "ORDER BY Parte;")
   If Not mRec.EOF Then
      mj = 1
      Do While Not mRec.EOF
         mj = mj + 1
         With FlexPartes
            .AddItem ""
            .TextMatrix(mj, 1) = mRec!Parte
            .TextMatrix(mj, 2) = NVL(mRec!FechaSolic, "")
            .TextMatrix(mj, 3) = NVL(mRec!CodEdificio, "")
            .TextMatrix(mj, 4) = NVL(mRec!descripcion, "")
            .TextMatrix(mj, 5) = NVL(mRec!Prioridad, "")
            .TextMatrix(mj, 6) = mRec!SectorAire
            .TextMatrix(mj, 7) = mRec!OpGen
         End With
         mRec.MoveNext
      Loop
   End If
   mRec.Close
End Sub

Private Sub cargarGrillaConPartesDeRelevamientos(ByVal pDescRamal As String, ByVal plistaPartesSeleccionados As String)
   'IMPORTANTE: El parametro plistaPartesSeleccionados no puede venir vacio porque da error, en tal caso se lo fuerza con el parte = -1
   Dim mj As Integer
   Dim sSql As String
'''Backup sentencia igual a la siguiente (por las dudas)
'''sSql = " SELECT * FROM ( " & _
'''                     " SELECT DISTINCT RD.IdRele,R.Parte,FechaSolic,CodEdificio,R.Descripcion,Prioridad,R.SectorAire, FechaIniAsist, concat(U.nombres, ' ',U.apellido ) as OpGen, R.CodSuperv   " & _
'''                     " FROM Registros R " & _
'''                        " Inner Join MailsxElectrico M ON R.SectorAire = M.SectorAire " & _
'''                        " Left Join OT_Partes OT ON OT.Parte = R.Parte " & _
'''                        " Inner Join REL_Relevamientos_Det RD ON RD.Parte = R.Parte " & _
'''                        " Inner Join REL_Relevamientos_H RH ON RD.IdRele = RH.IdRele " & _
'''                        " Left Join Cancelaciones_Det CNL ON CNL.Parte = R.Parte " & _
'''                        " Left Join loguser.usuarios U ON U.codusuario = R.OpGen " & _
'''                     " WHERE Estado NOT IN ('A', 'T') " & _
'''                     " AND Origen = 'O' AND M.codusuario = '" & Trim(Right(MDI.mUser, 20)) & "' " & _
'''                     " AND (OT.Parte IS NULL OR (OT.Cancelado = 1 " & _
'''                     " AND OT.Finalizado = 'NO' AND NOT EXISTS(SELECT 1 FROM OT_Partes WHERE Parte = OT.Parte AND Cancelado = 0))) " & _
'''                     " AND CodEdificio = '" & pDescRamal & "' " & _
'''                     " AND CNL.Parte IS NULL "
'''   sSql = sSql & " UNION "
'''   sSql = sSql & " SELECT DISTINCT RD.IdRele,R.Parte,FechaSolic,CodEdificio,R.descripcion,Prioridad,R.SectorAire, FechaIniAsist, concat(U.nombres, ' ',U.apellido ) as OpGen, R.CodSuperv   " & _
'''                     " FROM Registros R " & _
'''                        " Inner Join MailsxElectrico M ON R.SectorAire = M.SectorAire " & _
'''                        " Left Join OT_Partes OT ON OT.Parte = R.Parte " & _
'''                        " Inner Join REL_Relevamientos_Det_Columnas RD ON RD.Parte = R.Parte " & _
'''                        " Inner Join REL_Relevamientos_H RH ON RD.IdRele = RH.IdRele " & _
'''                        " Inner Join COM_Ramales CM ON CM.Codigo = RH.CodRamal " & _
'''                        " Left Join Cancelaciones_Det CNL ON CNL.Parte = R.Parte " & _
'''                        " Left Join loguser.usuarios U ON U.codusuario = R.OpGen " & _
'''                     " WHERE Estado NOT IN ('A', 'T') " & _
'''                     " AND Origen = 'O' AND M.codusuario = '" & Trim(Right(MDI.mUser, 20)) & "' " & _
'''                     " AND (OT.Parte IS NULL OR (OT.Cancelado = 1 " & _
'''                     " AND OT.Finalizado = 'NO' AND NOT EXISTS(SELECT 1 FROM OT_Partes WHERE Parte = OT.Parte AND Cancelado = 0))) " & _
'''                     " AND CM.Descripcion ='" & pDescRamal & "'" & _
'''                     " AND CNL.Parte IS NULL " & _
'''               " ) AUX WHERE AUX.Parte NOT IN (" & plistaPartesSeleccionados & ") " & _
'''                "ORDER BY Parte;"

   sSql = " SELECT * FROM ( " & _
                     " SELECT DISTINCT RD.IdRele,R.Parte,FechaSolic,CodEdificio,R.Descripcion,Prioridad,R.SectorAire, FechaIniAsist, concat(U.nombres, ' ',U.apellido ) as OpGen, R.CodSuperv   " & _
                     " FROM Registros R " & _
                        " Inner Join MailsxElectrico M ON R.SectorAire = M.SectorAire " & _
                        " Left Join OT_Partes OT ON OT.Parte = R.Parte " & _
                        " Inner Join REL_Relevamientos_Det RD ON RD.Parte = R.Parte " & _
                        " Inner Join REL_Relevamientos_H RH ON RD.IdRele = RH.IdRele " & _
                        " Left Join Cancelaciones_Det CNL ON CNL.Parte = R.Parte " & _
                        " Left Join loguser.usuarios U ON U.codusuario = R.OpGen " & _
                     " WHERE Estado NOT IN ('A', 'T') " & _
                     " AND Origen = 'O' AND M.codusuario = '" & Trim(Right(MDI.mUser, 20)) & "' " & _
                     " AND (OT.Parte IS NULL OR (OT.Cancelado = 1 " & _
                     " AND OT.Finalizado = 'NO' AND NOT EXISTS(SELECT 1 FROM OT_Partes WHERE Parte = OT.Parte AND Cancelado = 0))) " & _
                     " AND CodEdificio = '" & pDescRamal & "' " & _
                     " AND CNL.Parte IS NULL "
   sSql = sSql & " UNION "
   sSql = sSql & " SELECT DISTINCT RD.IdRele,R.Parte,FechaSolic,CodEdificio,R.descripcion,Prioridad,R.SectorAire, FechaIniAsist, concat(U.nombres, ' ',U.apellido ) as OpGen, R.CodSuperv   " & _
                     " FROM Registros R " & _
                        " Inner Join MailsxElectrico M ON R.SectorAire = M.SectorAire " & _
                        " Left Join OT_Partes OT ON OT.Parte = R.Parte " & _
                        " Inner Join REL_Relevamientos_Det_Columnas RD ON RD.Parte = R.Parte " & _
                        " Inner Join REL_Relevamientos_H RH ON RD.IdRele = RH.IdRele " & _
                        " Inner Join COM_Ramales CM ON CM.Codigo = RH.CodRamal " & _
                        " Left Join Cancelaciones_Det CNL ON CNL.Parte = R.Parte " & _
                        " Left Join loguser.usuarios U ON U.codusuario = R.OpGen " & _
                     " WHERE Estado NOT IN ('A', 'T') " & _
                     " AND Origen = 'O' AND M.codusuario = '" & Trim(Right(MDI.mUser, 20)) & "' " & _
                     " AND (OT.Parte IS NULL OR (OT.Cancelado = 1 " & _
                     " AND OT.Finalizado = 'NO' AND NOT EXISTS(SELECT 1 FROM OT_Partes WHERE Parte = OT.Parte AND Cancelado = 0))) " & _
                     " AND CM.Descripcion ='" & pDescRamal & "'" & _
                     " AND CNL.Parte IS NULL " & _
               " ) AUX WHERE AUX.Parte NOT IN (" & plistaPartesSeleccionados & ") " & _
                "ORDER BY Parte;"
   
   Set mRec = mObj.oEjecutarSelect(sSql)
   
   If Not mRec.EOF Then
      mj = 1
      Do While Not mRec.EOF
         mj = mj + 1
         With FlexPartes
            .AddItem ""
            .TextMatrix(mj, 1) = mRec!Parte
            .TextMatrix(mj, 2) = NVL(mRec!FechaSolic, "")
            .TextMatrix(mj, 3) = NVL(mRec!CodEdificio, "")
            .TextMatrix(mj, 4) = NVL(mRec!descripcion, "")
            .TextMatrix(mj, 5) = NVL(mRec!Prioridad, "")
            .TextMatrix(mj, 6) = mRec!SectorAire
            .TextMatrix(mj, 7) = mRec!OpGen
         End With
         mRec.MoveNext
      Loop
   End If
   mRec.Close
End Sub


