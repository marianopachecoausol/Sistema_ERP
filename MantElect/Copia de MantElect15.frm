VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form MantElect15 
   Caption         =   "Reposición de Materiales en Vehículos"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16965
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9405
   ScaleWidth      =   16965
   Begin VB.Frame Frame2 
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
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   16680
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   13560
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Última Reposición del Vehículo:"
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
         Left            =   10560
         TabIndex        =   11
         Top             =   270
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Reponer desde:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Reponer en Vehículo:"
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
         Left            =   5280
         TabIndex        =   8
         Top             =   270
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Confirmar"
      Height          =   305
      Index           =   0
      Left            =   5880
      TabIndex        =   4
      Top             =   8950
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   305
      Index           =   1
      Left            =   9480
      TabIndex        =   3
      Top             =   8950
      Width           =   1815
   End
   Begin VB.Frame Frame10 
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
      Height          =   8160
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   16680
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   278
         Left            =   14040
         TabIndex        =   12
         Top             =   8160
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   278
         Left            =   12960
         TabIndex        =   1
         Top             =   8155
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid FlexReposicion 
         Height          =   7800
         Left            =   120
         TabIndex        =   2
         Top             =   210
         Width           =   16440
         _ExtentX        =   28998
         _ExtentY        =   13758
         _Version        =   327680
         Cols            =   9
      End
   End
End
Attribute VB_Name = "MantElect15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mObj As New clMantElect
Dim mRec As New ADODB.Recordset

Dim mRenglonReposicion As Integer
Dim mCodProducto As String

Dim filaAnt As Integer
Dim columnAnt As Integer

Dim filaAntReponer As Integer
Dim columnAntReponer As Integer

Dim mvRepo_CodProd() As String
Dim mvRepo_StVeh_Inv() As Double
Dim mvRepo_StVeh_Sist() As Double
Dim mvRepo_Cant_Repuesta() As Double
Dim mvRepo_StDepo() As Double
Dim mvRepo_RepoSugerida_xMR() As Double

Dim mHayError As Boolean




Dim cboRamalListIndex As Integer

Private Sub Combo1_Click()
   '   Dim mi As Integer
   '   If cboRamalListIndex <> Combo1.ListIndex Then
   '      If cboRamalListIndex <> -1 Then
   '         If MsgBox("Si selecciona otro Ramal se perderán los datos cargados hasta el momento. ¿ Desea continuar ? ", vbYesNo, "Cambio de Ramal") = vbYes Then
   '            limpiarFormularioParcial
   '            'Elimino los registros grilla inferior
   '            For mi = FlexPartes.Rows To 3 Step -1
   '               FlexPartes.RemoveItem mi
   '            Next
   '         Else
   '            Combo1.ListIndex = cboRamalListIndex
   '         End If
   '         cboRamalListIndex = Combo1.ListIndex
   '      Else
   '         cboRamalListIndex = Combo1.ListIndex
   '      End If
   '   End If
   
   Dim sCodUbiDepo As String
   Dim sCodUbiVehiculo As String
   
   mHayError = False
   sCodUbiDepo = Right(Combo2.Text, 4)
   sCodUbiVehiculo = Right(Combo1.Text, 4)
   
   Set mRec = mObj.oEjecutarSelect(" SELECT IFNULL(Date_Format(MAX(Fecha),'%d-%m-%Y %h:%m:%s'),'') as Fecha FROM Reposiciones_H WHERE CodUbicacion_Destino = '" & sCodUbiVehiculo & "'; ")
   Text1.Text = mRec!Fecha
   
   cargarGrillaReposicion sCodUbiDepo, sCodUbiVehiculo
   preparaArrayReposiciones mvRepo_CodProd(), mvRepo_StVeh_Inv(), mvRepo_StVeh_Sist(), mvRepo_Cant_Repuesta(), mvRepo_StDepo(), mvRepo_RepoSugerida_xMR()
   
   mRenglonReposicion = 0
   Text2.Text = ""
   Text2.Visible = False
   Text3.Text = ""
   Text3.Visible = False
   FlexReposicion.ScrollBars = flexScrollBarVertical
End Sub

Private Sub cargarGrillaReposicion(ByVal pCodUbiDepo As String, ByVal pCodUbiVehiculo As String)
   Dim mj As Integer
   
   'Elimino los registros grilla inferior
   For mj = FlexReposicion.Rows To 3 Step -1
      FlexReposicion.RemoveItem mj
   Next
   
   Set mRec = mObj.oEjecutarSelect("   SELECT AUX.CodProducto, P.Descripcion AS Producto, UM.Descripcion AS UnidadMedida, AUX.CodUbicacion,SUM(AUX.Cantidad) AS RepoSugerida_x_MR ,SUM(AUX.StockVehiculo) AS StockVehiculo, SUM(AUX.StockDeposito) AS StockDeposito " & _
                                    "  FROM ( SELECT MR.CodProducto,MR.CodUbicacion,MR.Cantidad, M.Stock AS StockVehiculo, 0 As StockDeposito FROM " & _
                                    "     Matriz_Reposicion_Ubicaciones MR " & _
                                    "     LEFT JOIN Inventario.Movimientos2 M ON MR.CodUbicacion = M.CodUbicacion AND MR.CodProducto = M.CodProducto " & _
                                    "     WHERE MR.CodUbicacion = '" & pCodUbiVehiculo & "' " & _
                                    "     AND M.Fecha = (SELECT MAX(Fecha) From Inventario.Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion) " & _
                                    "        Union All " & _
                                    "     SELECT MR.CodProducto, MR.CodUbicacion,MR.Cantidad, 0 AS StockDestino, 0 As StockOigen FROM " & _
                                    "     Matriz_Reposicion_Ubicaciones MR " & _
                                    "     LEFT JOIN  Inventario.Movimientos2 M ON MR.CodUbicacion = M.CodUbicacion AND MR.CodProducto = M.CodProducto " & _
                                    "     WHERE MR.CodUbicacion = '" & pCodUbiVehiculo & "' " & _
                                    "     AND M.CodProducto IS NULL " & _
                                    "        Union All " & _
                                    "     SELECT MR.CodProducto,MR.CodUbicacion,0 As Cantidad, 0 as StockVehiculo,  IFNULL(M.Stock,0) AS StockDeposito FROM " & _
                                    "     Matriz_Reposicion_Ubicaciones MR " & _
                                    "     LEFT JOIN Inventario.Movimientos2 M ON M.CodUbicacion = '" & pCodUbiDepo & "' AND MR.CodProducto = M.CodProducto " & _
                                    "     AND M.Fecha = (  SELECT MAX(Fecha) From Inventario.Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion) " & _
                                    "     WHERE MR.CodUbicacion = '" & pCodUbiVehiculo & "') AS AUX " & _
                                    " INNER JOIN Inventario.Producto P ON P.Codigo = AUX.CodProducto " & _
                                    " INNER JOIN Inventario.UnidadMedida UM ON UM.Codigo = P.CodUnidadMedida " & _
                                    "     GROUP BY AUX.CodProducto, P.Descripcion, UM.Descripcion , AUX.CodUbicacion; ")
                                    
'SELECT AUX.CodProducto, P.Descripcion AS Producto, UM.Descripcion AS UnidadMedida, AUX.CodUbicacion,SUM(AUX.Cantidad) AS CantReponer ,SUM(AUX.CantidadBD) AS CantReponerBD, SUM(AUX.StockVehiculo) AS StockVehiculo, SUM(AUX.StockDeposito) AS StockDeposito
'FROM ( SELECT MR.CodProducto,MR.CodUbicacion,MR.Cantidad, MR.Cantidad AS CantidadBD ,M.Stock AS StockVehiculo, 0 As StockDeposito FROM
       'Matriz_Reposicion_Ubicaciones MR
       'LEFT JOIN Inventario.Movimientos2 M ON MR.CodUbicacion = M.CodUbicacion AND MR.CodProducto = M.CodProducto
       'WHERE MR.CodUbicacion = '0011'
       'AND M.Fecha = (SELECT MAX(Fecha) From Inventario.Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion)
         'Union All
       'SELECT MR.CodProducto, MR.CodUbicacion,MR.Cantidad, MR.Cantidad AS CantidadBD, 0 AS StockDestino, 0 As StockOigen FROM
       'Matriz_Reposicion_Ubicaciones MR
       'LEFT JOIN  Inventario.Movimientos2 M ON MR.CodUbicacion = M.CodUbicacion AND MR.CodProducto = M.CodProducto
       'WHERE MR.CodUbicacion = '0011'
       'AND M.CodProducto IS NULL
        ' Union All
       'SELECT MR.CodProducto,MR.CodUbicacion,0 As Cantidad,0 as CantidadBD, 0 as StockVehiculo,  IFNULL(M.Stock,0) AS StockDeposito FROM
       'Matriz_Reposicion_Ubicaciones MR
       'LEFT JOIN Inventario.Movimientos2 M ON M.CodUbicacion = '0010' AND MR.CodProducto = M.CodProducto
       'AND M.Fecha = (  SELECT MAX(Fecha) From Inventario.Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion)
       'WHERE MR.CodUbicacion = '0011') AS AUX
'INNER JOIN Inventario.Producto P ON P.Codigo = AUX.CodProducto
'INNER JOIN Inventario.UnidadMedida UM ON UM.Codigo = P.CodUnidadMedida
       'GROUP BY AUX.CodProducto, P.Descripcion, UM.Descripcion , AUX.CodUbicacion;
                                
   If Not mRec.EOF Then
      mj = 1
      Do While Not mRec.EOF
         mj = mj + 1
         With FlexReposicion
            .AddItem ""
            .TextMatrix(mj, 1) = mRec!CodProducto
            .TextMatrix(mj, 2) = mRec!Producto
            .TextMatrix(mj, 3) = ""
            .TextMatrix(mj, 4) = mRec!StockVehiculo
            .TextMatrix(mj, 5) = ""
            .TextMatrix(mj, 6) = mRec!StockDeposito
            .TextMatrix(mj, 7) = mRec!UnidadMedida
            .TextMatrix(mj, 8) = mRec!RepoSugerida_x_MR
         End With
         mRec.MoveNext
      Loop
   End If
   mRec.Close
   
End Sub


Private Sub preparaArrayReposiciones(ByRef pvRepo_CodProd() As String, ByRef pvRepo_StVeh_Inv() As Double, ByRef pvRepo_StVeh_Sist() As Double, _
                                       ByRef pvRepo_Cant_Repuesta() As Double, ByRef pvRepo_StDepo() As Double, ByRef pvRepo_RepoSugerida_xMR() As Double)
   Dim mj As Integer
   Dim cantMateriales As Integer

   cantMateriales = FlexReposicion.Rows - 2
   If cantMateriales > 0 Then

      
      ReDim pvRepo_CodProd(0 To cantMateriales - 1) As String
      ReDim pvRepo_StVeh_Inv(0 To cantMateriales - 1) As Double
      ReDim pvRepo_StVeh_Sist(0 To cantMateriales - 1) As Double
      ReDim pvRepo_Cant_Repuesta(0 To cantMateriales - 1) As Double
      ReDim pvRepo_StDepo(0 To cantMateriales - 1) As Double
      ReDim pvRepo_RepoSugerida_xMR(0 To cantMateriales - 1) As Double
      
      For mj = 2 To FlexReposicion.Rows - 1
         pvRepo_CodProd(mj - 2) = FlexReposicion.TextMatrix(mj, 1)
         
         If FlexReposicion.TextMatrix(mj, 3) = "" Then
            pvRepo_StVeh_Inv(mj - 2) = 0
         Else
            pvRepo_StVeh_Inv(mj - 2) = CDbl(Replace(FlexReposicion.TextMatrix(mj, 3), ".", ","))
         End If
        
         pvRepo_StVeh_Sist(mj - 2) = CDbl(Replace(FlexReposicion.TextMatrix(mj, 4), ".", ","))
        
         If FlexReposicion.TextMatrix(mj, 5) = "" Then
            pvRepo_Cant_Repuesta(mj - 2) = 0
         Else
            pvRepo_Cant_Repuesta(mj - 2) = CDbl(Replace(FlexReposicion.TextMatrix(mj, 5), ".", ","))
         End If
        
        pvRepo_StDepo(mj - 2) = CDbl(Replace(FlexReposicion.TextMatrix(mj, 6), ".", ","))
        pvRepo_RepoSugerida_xMR(mj - 2) = CDbl(Replace(FlexReposicion.TextMatrix(mj, 8), ".", ","))
      Next
   Else
      'Esta linea sirve como flag para avisar en el procedimiento xinsOT que no tiene registros y por ende no me de error de intervalos al usar Ubound y Lubound
      ReDim pvRepo_CodProd(0)
      pvRepo_CodProd(0) = "000000"
   End If
End Sub




Private Sub Command2_Click(Index As Integer)

If Index = 0 Then
   Dim sCodUbiDepo As String
   Dim sCodUbiVehiculo As String
   
   sCodUbiDepo = Trim(Right(Combo2.Text, 4))
   sCodUbiVehiculo = Trim(Right(Combo1.Text, 4))
   
   If Not mHayError Then
      If fValidaConfirmarRepoGrilla() Then
         preparaArrayReposiciones mvRepo_CodProd(), mvRepo_StVeh_Inv(), mvRepo_StVeh_Sist(), mvRepo_Cant_Repuesta(), mvRepo_StDepo(), mvRepo_RepoSugerida_xMR()
         
         If fValidaConfirmarRepoArray(mvRepo_CodProd(), mvRepo_Cant_Repuesta(), sCodUbiDepo) Then
            sMsgEspere Me, "Procesando reposición...", True
            mObj.xinsRepo Trim(Right(MDI.mUser, 15)), sCodUbiDepo, sCodUbiVehiculo, mvRepo_CodProd(), mvRepo_StVeh_Inv(), mvRepo_StVeh_Sist(), mvRepo_Cant_Repuesta(), mvRepo_StDepo(), mvRepo_RepoSugerida_xMR
            sMsgEspere Me, "", False
            
            MsgBox "Se ha realizado exitosamente la reposicion en el vehiculo seleccionado ", vbInformation
            
            sLlenoComboVehiculoDestino
            sCodUbiDepo = Trim(Right(Combo2.Text, 4))
            sCodUbiVehiculo = Trim(Right(Combo1.Text, 4))
            
            Text2.Visible = False
            Text2.Text = ""
            Text3.Visible = False
            Text3.Text = ""
            
            Text1.Text = ""
            
            cargarGrillaReposicion sCodUbiDepo, sCodUbiVehiculo
         Else
            Exit Sub
         End If
      Else
         Exit Sub
      End If
   End If
Else
   Unload Me
End If

End Sub

''''Private Sub preparaArrayPartes(ByRef pvParte_CodActivo() As String, ByRef pvParte_DescActivo() As String, ByRef pvParte_CodProblema() As String, ByRef pvParte_DescProblema() As String, ByRef pvParte_Prioridad() As String)
''''   Dim mj As Integer
''''   Dim cantPartes As Integer
''''
''''  cantPartes = FlexPartes.Rows - 2
''''   If cantPartes > 0 Then
''''
''''      ReDim pvParte_CodActivo(0 To cantPartes - 1) As String
''''      ReDim pvParte_DescActivo(0 To cantPartes - 1) As String
''''      ReDim pvParte_CodProblema(0 To cantPartes - 1) As String
''''      ReDim pvParte_DescProblema(0 To cantPartes - 1) As String
''''      ReDim pvParte_Prioridad(0 To cantPartes - 1) As String
''''
''''      For mj = 2 To FlexPartes.Rows - 1
''''        pvParte_CodActivo(mj - 2) = FlexPartes.TextMatrix(mj, 4)
''''        pvParte_DescActivo(mj - 2) = FlexPartes.TextMatrix(mj, 1)
''''        pvParte_CodProblema(mj - 2) = FlexPartes.TextMatrix(mj, 5)
''''        pvParte_DescProblema(mj - 2) = FlexPartes.TextMatrix(mj, 2)
''''        pvParte_Prioridad(mj - 2) = FlexPartes.TextMatrix(mj, 3)
''''      Next
''''   End If
''''End Sub


''''Private Function fValidaComunicado() As Boolean
''''   Dim mRet As Boolean
''''   Dim mMensajeError As String
''''   Dim mRec1 As New ADODB.Recordset
''''
''''   mRet = True
''''
''''   If Trim(Text1.Text) = "" Then
''''      mRet = False
''''      mMensajeError = "Debe completar el campo 'Comunicado'"
''''   End If
''''
''''   If mRet Then
''''      Set mRec1 = mObj.oEjecutarSelect("SELECT * FROM COM_Comunicados_H WHERE NroComunicado = '" & Trim(Text1.Text) & "';")
''''      If Not mRec1.EOF Then
''''         mRet = False
''''         mMensajeError = "Ya existe un comunicado con ese número."
''''      End If
''''      mRec1.Close
''''   End If
''''
''''   If mRet Then
''''      If Trim(Combo1.Text) = "" Then
''''         mRet = False
''''         mMensajeError = "Debe seleccionar un Ramal"
''''      End If
''''   End If
''''
''''   If mRet Then
''''      If FlexPartes.Rows <= 2 Then
''''         mRet = False
''''         mMensajeError = "Al menos se debe crear un Parte"
''''      End If
''''   End If
''''
''''   Set mRec1 = Nothing
''''
''''   If Not mRet Then
''''         MsgBox mMensajeError, vbCritical, "Atención"
''''   End If
''''
''''   fValidaComunicado = mRet
''''End Function

''''Private Sub limpiarFormularioTotal()
''''   Dim mi As Integer
''''   Text1.Text = ""
''''   'sLlenoRamal
''''   limpiarFormularioParcial
''''   'Elimino los registros grilla inferior
''''   For mi = FlexPartes.Rows To 3 Step -1
''''      FlexPartes.RemoveItem mi
''''   Next
''''End Sub

''''Private Sub FlexPartes_Click()
''''   If FlexPartes.MouseCol = 0 And FlexPartes.MouseRow > 0 Then
''''      If FlexPartes.Rows > 2 Then
''''         FlexPartes.RemoveItem FlexPartes.MouseRow
''''         FlexPartes.Row = 0
''''      End If
''''   End If
''''End Sub


Private Sub FlexReposicion_Click()
Dim mi As Integer
   
If Not mHayError Then
   If FlexReposicion.MouseRow > 0 Then

         'En este caso 3 es la columna que seria editable
         If FlexReposicion.Col = 3 And FlexReposicion.Row <> 1 Then
            Text2.Text = FlexReposicion.Text
            Text2.Width = FlexReposicion.ColWidth(FlexReposicion.Col)
            Text2.Left = FlexReposicion.ColPos(FlexReposicion.Col) + FlexReposicion.Left + 30 'el valor treina termina de acomodar el textbox en la celda
            Text2.Top = FlexReposicion.Top + FlexReposicion.RowPos(FlexReposicion.Row)
            Text2.Visible = True
            Text2.SetFocus
            
            FlexReposicion.ScrollBars = flexScrollBarNone
         Else
            Text2.Visible = False
            FlexReposicion.ScrollBars = flexScrollBarVertical
         End If
         filaAnt = FlexReposicion.Row
         columnAnt = FlexReposicion.Col
         'En este caso 3 es la columna que seria editable
         If FlexReposicion.Col = 5 And FlexReposicion.Row <> 1 Then
            Text3.Text = FlexReposicion.Text
            Text3.Width = FlexReposicion.ColWidth(FlexReposicion.Col)
            Text3.Left = FlexReposicion.ColPos(FlexReposicion.Col) + FlexReposicion.Left + 30 'el valor treina termina de acomodar el textbox en la celda
            Text3.Top = FlexReposicion.Top + FlexReposicion.RowPos(FlexReposicion.Row)
            Text3.Visible = True
            Text3.SetFocus

            FlexReposicion.ScrollBars = flexScrollBarNone
         Else
            Text3.Visible = False
            FlexReposicion.ScrollBars = flexScrollBarVertical
         End If
         filaAntReponer = FlexReposicion.Row
         columnAntReponer = FlexReposicion.Col

      If mRenglonReposicion <> 0 Then
         FlexReposicion.Row = mRenglonReposicion
         For mi = 1 To FlexReposicion.Cols - 1
            FlexReposicion.Col = mi
            FlexReposicion.CellBackColor = vbWhite
         Next
      End If
      mRenglonReposicion = FlexReposicion.MouseRow
      FlexReposicion.Row = mRenglonReposicion
      For mi = 1 To FlexReposicion.Cols - 1
         FlexReposicion.Col = mi
         FlexReposicion.CellBackColor = &H80000003
      Next
      If mRenglonReposicion > 1 Then
          mCodProducto = FlexReposicion.TextMatrix(mRenglonReposicion, 1)
      End If
   Else
      FlexReposicion.Row = mRenglonReposicion
      If FlexReposicion.Row > 0 Then
         For mi = 1 To FlexReposicion.Cols - 1
            FlexReposicion.Col = mi
            FlexReposicion.CellBackColor = vbWhite
         Next
      End If
      mRenglonReposicion = 0
   End If
End If
End Sub

Private Sub Form_Load()
   Me.Width = 17085
   Me.Height = 9920
   sAlinearForm Me

   mHayError = False

   sLlenoComboDesde
   Combo2.ListIndex = 0
   sLlenoComboVehiculoDestino
   initReposicion
End Sub

Private Sub initReposicion()
   filaAnt = 0
   columnAnt = 0
   filaAntReponer = 0
   columnAntReponer = 0
   Text2.Visible = False
   Text3.Visible = False
   
   With FlexReposicion
      .ColWidth(0) = 200
      .ColWidth(1) = 900
      .ColWidth(2) = 8120
      .ColWidth(3) = 1500
      .ColWidth(4) = 1500
      .ColWidth(5) = 1500
      .ColWidth(6) = 1500
      .ColWidth(7) = 850
      .ColWidth(8) = 0

      .TextMatrix(0, 1) = "Cód.Prod."
      .TextMatrix(0, 2) = "Producto"
      .TextMatrix(0, 3) = "St.Vehíc.Inventario"
      .TextMatrix(0, 4) = "St.Vehíc.Sistema"
      .TextMatrix(0, 5) = "Reponer"
      .TextMatrix(0, 6) = "St.Depósito"
      .TextMatrix(0, 7) = "U.Medida"
      .TextMatrix(0, 8) = "RepoSugerida_xMR"
      
      .RowHeight(1) = 0
   End With
End Sub


Private Sub sLlenoComboDesde()
   
   Set mRec = mObj.oEjecutarSelect(" SELECT U.Codigo,U.Descripcion,U.CodBodega, V.CodUbicacion " & _
                                 " FROM " & _
                                 "   Inventario.Ubicaciones U " & _
                                 " INNER JOIN " & _
                                 "   Inventario.Usuario_AccesoBodega AB ON U.CodBodega = AB.CodBodega " & _
                                 " LEFT JOIN " & _
                                 "   MantElect.Vehiculos V ON V.CodUbicacion = U.Codigo " & _
                                 " Where V.CodUbicacion Is Null " & _
                                 " AND  AB.codusuario = '" & Trim(Right(MDI.mUser, 15)) & "' " & _
                                 " AND U.Fecha_Baja IS NULL; ")
                                 
   Do While Not mRec.EOF
      Combo2.AddItem "" & mRec!descripcion & Space(80) & mRec!Codigo & ""
      mRec.MoveNext
   Loop
   mRec.Close
End Sub

Private Sub sLlenoComboVehiculoDestino()
   Combo1.Clear
   Set mRec = mObj.oEjecutarSelect(" SELECT U.Codigo,U.Descripcion,U.CodBodega, V.CodUbicacion " & _
                                 " FROM " & _
                                 "   Inventario.Ubicaciones U " & _
                                 " INNER JOIN " & _
                                 "   Inventario.Usuario_AccesoBodega AB ON U.CodBodega = AB.CodBodega " & _
                                 " LEFT JOIN " & _
                                 "   MantElect.Vehiculos V ON V.CodUbicacion = U.Codigo " & _
                                 " Where V.CodUbicacion Is Not Null " & _
                                 " AND  AB.codusuario = '" & Trim(Right(MDI.mUser, 15)) & "' " & _
                                 " AND U.Fecha_Baja IS NULL; ")
                                 
   Do While Not mRec.EOF
      Combo1.AddItem "" & mRec!descripcion & Space(80) & mRec!Codigo & ""
      mRec.MoveNext
   Loop
   mRec.Close
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
      If fValidaCantidad(Text2.Text, True) Then
         FlexReposicion.TextMatrix(filaAnt, columnAnt) = Text2.Text
         Text2.Visible = False
         FlexReposicion.ScrollBars = flexScrollBarVertical
      End If
   End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 46 Then
      KeyAscii = fNumeroKeyPress(KeyAscii)
   End If
   
   If KeyAscii = 13 Then
      If fValidaCantidad(Text3.Text, True) Then
         FlexReposicion.TextMatrix(filaAntReponer, columnAntReponer) = Text3.Text
         Text3.Visible = False
         FlexReposicion.ScrollBars = flexScrollBarVertical
      End If
   End If
End Sub


Private Sub Text2_LostFocus()
   If FlexReposicion.Col <> columnAnt Or FlexReposicion.Row <> filaAnt Then
      'En este caso 3 es la columna que seria editable
      If columnAnt = 3 Then
         If fValidaCantidad(Text2.Text, True) Then
            FlexReposicion.TextMatrix(filaAnt, columnAnt) = Text2.Text
               If CDbl(Replace(Trim(FlexReposicion.TextMatrix(filaAnt, columnAnt)), ".", ",")) <= CDbl(Replace(Trim(FlexReposicion.TextMatrix(filaAnt, columnAnt + 5)), ".", ",")) Then
                  FlexReposicion.TextMatrix(filaAnt, columnAnt + 2) = CDbl(Replace(Trim(FlexReposicion.TextMatrix(filaAnt, columnAnt + 5)), ".", ",")) - CDbl(Replace(Trim(FlexReposicion.TextMatrix(filaAnt, columnAnt)), ".", ","))
                  FlexReposicion.TextMatrix(filaAnt, columnAnt + 2) = Replace(Trim(FlexReposicion.TextMatrix(filaAnt, columnAnt + 2)), ",", ".")
               Else
                  FlexReposicion.TextMatrix(filaAnt, columnAnt + 2) = 0
               End If
         End If
      End If
   End If
End Sub

Private Sub Text3_LostFocus()
   If FlexReposicion.Col <> columnAntReponer Or FlexReposicion.Row <> filaAntReponer Then
      'En este caso 5 es la columna que seria editable
      If columnAntReponer = 5 Then
         If fValidaCantidad(Text3.Text, True) Then
            FlexReposicion.TextMatrix(filaAntReponer, columnAntReponer) = Text3.Text
         End If
      End If
   End If
End Sub

Private Function fValidaCantidad(ByVal pCantidad As String, ByVal pMuestraMsgError As Boolean) As Boolean
   Dim mRet As Boolean
   Dim mMensajeError As String
   Dim posInstr As Integer
   Dim qtyDecimales As Integer
   
   mRet = True
      
   If Trim(pCantidad) = "" Then
      mRet = False
      mMensajeError = "Debe completar con un valor numérico. "
   End If
      
   If mRet Then
      If Not IsNumeric(Replace(pCantidad, ".", ",")) Then
         mRet = False
         mMensajeError = "El valor ingresado no es numérico"
      End If
   End If
   
   If mRet Then
      If CDbl(Replace(Trim(pCantidad), ".", ",")) < 0 Then
         mRet = False
         mMensajeError = "La cantidad ingresada no puede ser menor o igual a cero."
      End If
   End If
   
   'Valido que no supere los 2 digitos decimales
   If mRet Then
      posInstr = InStr(1, Replace(Trim(pCantidad), ".", ","), ",")

      If posInstr <> 0 Then
         qtyDecimales = Len(Right(Trim(pCantidad), Len(Trim(pCantidad)) - posInstr))
      End If

      If qtyDecimales > 2 Then
         mRet = False
         mMensajeError = "Solo se admiten hasta dos dígitos decimales."
      End If
   End If
   
   mHayError = Not mRet
   
   If Not mRet Then
      If pMuestraMsgError Then
         MsgBox mMensajeError, vbCritical, "Atención"
      End If
   End If
   fValidaCantidad = mRet

End Function


'fValidaConfirmarAjusteArray(mvRepo_CodProd(), mvRepo_Cant_Repuesta(), sCodUbiDepo)
Public Function fValidaConfirmarRepoArray(ByRef pvRepo_CodProd() As String, ByRef pvRepo_Cant_Repuesta() As Double, ByVal pCodUbiDepo As String) As Boolean
   Dim mi As Integer
   Dim mRet As Boolean
   Dim mMensajeError As String
   Dim mRec1 As New ADODB.Recordset
   Dim iStock As Double
   Dim ProdDescr As String
   
   mRet = True
   
   If pvRepo_CodProd(0) <> "000000" Then
      For mi = LBound(pvRepo_CodProd) To UBound(pvRepo_CodProd)
            Set mRec1 = mObj.oEjecutarSelect("SELECT Fecha,CodProducto,CodUbicacion, IFNULL(Stock,0) As Stock " & _
                                             " FROM Inventario.Movimientos2 M " & _
                                             " WHERE CodProducto  = '" & pvRepo_CodProd(mi) & "' and CodUbicacion = '" & pCodUbiDepo & "'" & _
                                             " AND Fecha = (SELECT Max(Fecha) FROM Inventario.Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion)")
            If Not mRec1.EOF Then
               iStock = mRec1!Stock
            Else
               iStock = 0
            End If
            mRec1.Close
            
            If pvRepo_Cant_Repuesta(mi) > iStock Then
               ProdDescr = mObj.sCampoDescrip("Inventario.Producto", "Codigo='" & pvRepo_CodProd(mi) & "'", 1)
               mRet = False
               mMensajeError = "El stock es insuficiente para ' " & ProdDescr & " '"
               mi = 9999
            End If
      Next
   End If
   
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   
   fValidaConfirmarRepoArray = mRet
End Function



Public Function fValidaConfirmarRepoGrilla() As Boolean

   Dim mRet As Boolean
   Dim mMensajeError As String
   Dim cantMateriales As Integer
   Dim mi As Integer
   Dim posInstr As Integer
   Dim qtyDecimales As Integer
   

    
   mRet = True

   cantMateriales = FlexReposicion.Rows - 2
   If cantMateriales <= 0 Then
      mRet = False
      mMensajeError = "La grilla debe contener materiales."
   End If
   

   If mRet Then
      For mi = 2 To FlexReposicion.Rows - 1
         '*************************************************************************************************************************************
         '-------------------------------------VALIDACION COLUMNA STOCK INVENTARIADO-----------------------------------------------------------
         '*************************************************************************************************************************************
        'Valido valor numerico
         If Not IsNumeric(Replace(FlexReposicion.TextMatrix(mi, 3), ".", ",")) Then
            mRet = False
            mMensajeError = "El valor de la columna 'St.Vehíc.Inventario' es incorrecto para el producto: '" & FlexReposicion.TextMatrix(mi, 2) & "'"
            mi = 9999
         End If
         
         'Valido cantidad decimales
         If mRet Then
            posInstr = InStr(1, Replace(FlexReposicion.TextMatrix(mi, 3), ".", ","), ",")
      
            qtyDecimales = 0
            If posInstr <> 0 Then
               qtyDecimales = Len(Right(Trim(FlexReposicion.TextMatrix(mi, 3)), Len(Trim(FlexReposicion.TextMatrix(mi, 3))) - posInstr))
            End If
            If qtyDecimales > 2 Then
               mRet = False
               mMensajeError = "El valor de la columna 'St.Vehíc.Inventario' no puede tener mas de dos decimales para el producto: '" & FlexReposicion.TextMatrix(mi, 2) & "'"
               mi = 9999
            End If
         End If
         
         'Valido valor > 0
         If mRet Then
            If CDbl(Replace(Trim(FlexReposicion.TextMatrix(mi, 3)), ".", ",")) < 0 Then
               mRet = False
               mMensajeError = "El valor de la columna 'St.Vehíc.Inventario' es menor a cero para el producto: '" & FlexReposicion.TextMatrix(mi, 2) & "'"
               mi = 9999
            End If
         End If
         '-------------------------------------FIN VALIDACION COLUMNA STOCK INVENTARIADO-------------------------------------------------------
         
         
         '*************************************************************************************************************************************
         '-------------------------------------VALIDACION COLUMNA REPONER----------------------------------------------------------------------
         '*************************************************************************************************************************************
         'Valido valor numerico
         If mRet Then
            If Not IsNumeric(Replace(FlexReposicion.TextMatrix(mi, 5), ".", ",")) Then
               mRet = False
               mMensajeError = "El valor de la columna 'Reponer' es incorrecto para el producto: '" & FlexReposicion.TextMatrix(mi, 2) & "'"
               mi = 9999
            End If
         End If
         
         'Valido cantidad decimales
         If mRet Then
            posInstr = InStr(1, Replace(FlexReposicion.TextMatrix(mi, 5), ".", ","), ",")
      
            qtyDecimales = 0
            If posInstr <> 0 Then
               qtyDecimales = Len(Right(Trim(FlexReposicion.TextMatrix(mi, 5)), Len(Trim(FlexReposicion.TextMatrix(mi, 5))) - posInstr))
            End If
            If qtyDecimales > 2 Then
               mRet = False
               mMensajeError = "El valor de la columna 'Reponer' no puede tener más de dos decimales para el producto: '" & FlexReposicion.TextMatrix(mi, 2) & "'"
               mi = 9999
            End If
         End If
         
         'Valido valor > 0
         If mRet Then
            If CDbl(Replace(Trim(FlexReposicion.TextMatrix(mi, 5)), ".", ",")) < 0 Then
               mRet = False
               mMensajeError = "El valor de la columna 'Reponer' es menor a cero para el producto: '" & FlexReposicion.TextMatrix(mi, 2) & "'"
               mi = 9999
            End If
         End If
         
         '-------------------------------------FIN VALIDACION COLUMNA REPONDER-------------------------------------------------------------------
      Next
   End If
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   fValidaConfirmarRepoGrilla = mRet
End Function



