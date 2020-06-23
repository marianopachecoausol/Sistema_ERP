VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Inven011_frmGrande 
   Caption         =   "Nuevo Ingreso de Materiales"
   ClientHeight    =   13440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   21150
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   13440
   ScaleWidth      =   21150
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   12280
      MouseIcon       =   "Inven011_frm.frx":0000
      TabIndex        =   17
      Top             =   12720
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Confirmar Ingreso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   6840
      MouseIcon       =   "Inven011_frm.frx":030A
      TabIndex        =   12
      Top             =   12720
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   0
      Left            =   3000
      Picture         =   "Inven011_frm.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   1
      Left            =   3500
      Picture         =   "Inven011_frm.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7440
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   8
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ingresos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   8040
      Width           =   20895
      Begin MSFlexGridLib.MSFlexGrid FlexIngreso 
         Height          =   3615
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   20415
         _ExtentX        =   36010
         _ExtentY        =   6376
         _Version        =   327680
         Cols            =   8
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selección del Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   20895
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13320
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   600
         Width           =   10455
      End
      Begin MSFlexGridLib.MSFlexGrid FlexProduct 
         Height          =   3735
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   20415
         _ExtentX        =   36010
         _ExtentY        =   6588
         _Version        =   327680
         Cols            =   8
      End
      Begin VB.Label Label3 
         Caption         =   "Contiene texto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información del Ingreso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   20895
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   15820
         MaxLength       =   10
         TabIndex        =   16
         Top             =   640
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   640
         Width           =   3735
      End
      Begin VB.Label Label5 
         Caption         =   "Orden de Compra:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13680
         TabIndex        =   15
         Top             =   700
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Ingresar en Bodega:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   700
         Width           =   2055
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Cantidad:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   7
      Top             =   7500
      Width           =   975
   End
End
Attribute VB_Name = "Inven011_frmGrande"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mObj As New clInven
Dim mRec As ADODB.Recordset
Dim mRenglonProducto As Integer
Dim mRenglonIngreso As Integer
Dim mCodProducto As String
Dim cboListIndex As Integer
'TODO: Reemplazar flexIngreso X FlexIngreso

Private Sub Combo1_Click()
   Dim mi As Integer
   
   If cboListIndex <> Combo1.ListIndex Then

      If (cboListIndex <> -1) Then
         
         If MsgBox("Si selecciona otra Bodega se perderán los ingresos cargados hasta el momento. ¿ Desea continuar ? ", vbYesNo, "Cambio de Bodega") = vbYes Then
            Text1.Text = ""
            Text2.Text = ""
            
            'Elimino los registros de la grilla superior (productos)
            For mi = FlexProduct.Rows To 3 Step -1
               FlexProduct.RemoveItem mi
            Next
            mRenglonProducto = 0
            
            'Elimino los registros de la grilla inferior (Ingresos)
            For mi = FlexIngreso.Rows To 3 Step -1
               FlexIngreso.RemoveItem mi
            Next
            mRenglonIngreso = 0
           
         Else
            Combo1.ListIndex = cboListIndex
         End If
         
         cboListIndex = Combo1.ListIndex
      
      Else
         cboListIndex = Combo1.ListIndex
      End If
   
   End If
End Sub

Private Sub Command1_Click()
   Dim mi As Integer
   Dim mJ As Integer
   
   mRenglonProducto = 0
   
   'Elimino los registros (de la consulta anterior) de la grilla superior
   For mi = FlexProduct.Rows To 3 Step -1
      FlexProduct.RemoveItem mi
   Next
   
   Set mRec = mObj.getStockXBodegaConFiltroProducto(Left(Combo1.Text, 4), Text1.Text)
   
   'Cargo la Grilla Superior con lo devuelto por el sp "getStockXBodegaConFiltroProducto"
   If Not mRec.EOF Then
      mi = 1
      Do While Not mRec.EOF
         mi = mi + 1
         
         FlexProduct.AddItem ""
         FlexProduct.TextMatrix(mi, 1) = mRec!Producto
         FlexProduct.TextMatrix(mi, 2) = mRec!Ubicacion
         FlexProduct.TextMatrix(mi, 3) = mRec!Stock
         FlexProduct.TextMatrix(mi, 4) = mRec!UnidadMedida
         FlexProduct.TextMatrix(mi, 5) = mRec!CodigoSap
         FlexProduct.TextMatrix(mi, 6) = mRec!CodProducto
         FlexProduct.TextMatrix(mi, 7) = mRec!CodUbicacion
         
         mRec.MoveNext
      Loop
   End If
   mRec.Close
   
    'COMENTO LAS LINEAS SIGUIENTES PORQUE CUANDO AJUSTO (NO SUMO NI RESTO EN LA GRILLA SUPERIOR)
'   'Si presiono el boton Buscar y algun "producto/ubicacion" de la grilla de arriba, esta en la grilla inferior
'   'entonces que me actualice en la grilla superior el stock de ese o esos "producto/ubicacion" agregando el ingreso de la grilla inferior
'   For mI = 2 To FlexProduct.Rows - 1
'      For mJ = 2 To FlexIngreso.Rows - 1
'         If FlexProduct.TextMatrix(mI, 6) = FlexIngreso.TextMatrix(mJ, 6) And FlexProduct.TextMatrix(mI, 7) = FlexIngreso.TextMatrix(mJ, 7) Then
'            FlexProduct.TextMatrix(mI, 3) = CDbl(Replace(Trim(FlexProduct.TextMatrix(mI, 3)), ".", ",")) + CDbl(Replace(Trim(FlexIngreso.TextMatrix(mJ, 3)), ".", ","))
'            mJ = 999
'         End If
'      Next
'   Next
End Sub

Private Sub Command2_Click(Index As Integer)
   Dim iStock As Double
   Dim mi As Integer
   Dim mRec1 As New ADODB.Recordset
   
   If Index = 0 Then
      'TODO(Realizado): Debe ser fValidaIngreso.
      'TODO(Realizado):Redefinir las validaciones en esta nueva fValidaIngreso
      If fValidaIngreso() Then
            FlexIngreso.AddItem vbTab & FlexProduct.TextMatrix(mRenglonProducto, 1) & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 2) & vbTab & Text2.Text & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 4) & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 5) & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 6) & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 7)
            'No sumo en la grilla superior, por eso comente la siguiente linea.
            'FlexProduct.TextMatrix(mRenglonProducto, 3) = CDbl(Replace(Trim(FlexProduct.TextMatrix(mRenglonProducto, 3)), ".", ",")) + CDbl(Replace(Trim(Text2.Text), ".", ","))
            Text2.Text = ""
            Text2.SetFocus
      End If
   Else
'Comento este For debido a que no es necesario restar en la grilla superior lo que habia elegido y luego devuelvo.
'      For mi = 2 To FlexProduct.Rows - 1
'
'         If FlexProduct.TextMatrix(mi, 6) = FlexIngreso.TextMatrix(mRenglonIngreso, 6) And FlexProduct.TextMatrix(mi, 7) = FlexIngreso.TextMatrix(mRenglonIngreso, 7) Then
'            Set mRec1 = mObj.oEjecutarSelect("SELECT Fecha,CodProducto,CodUbicacion, IFNULL(Stock,0) As Stock " & _
'                                       " FROM Movimientos2 M " & _
'                                       " WHERE CodProducto  = '" & FlexIngreso.TextMatrix(mRenglonIngreso, 6) & "' and CodUbicacion = '" & FlexIngreso.TextMatrix(mRenglonIngreso, 7) & "'" & _
'                                       " AND Fecha = (SELECT Max(Fecha) FROM Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion)")
'
'            If Not mRec1.EOF Then
'               iStock = mRec1!Stock
'            Else
'               iStock = 0
'            End If
'            mRec1.Close
'
'            FlexProduct.TextMatrix(mi, 3) = iStock
'
'            mi = 9999
'         End If
'      Next
      
      If FlexIngreso.Rows > 2 And mRenglonIngreso > 1 Then
         FlexIngreso.RemoveItem (mRenglonIngreso)
      End If
      
      mRenglonIngreso = 0
   End If
End Sub

'Boton de confirmacion de Ingreso de materiales o salir
Private Sub Command3_Click(Index As Integer)
   If Index = 0 Then
      Dim vIngresosCodProducto() As String
      Dim vIngresosCodUbicacion() As String
      Dim vIngresosCantidad() As Double
      'TODO(Realizado): Reemplazar variable cantEgresos por cantIngresos
      Dim cantIngresos As Integer
      Dim mJ As Integer
      Dim mCodTipoVale As String
      Dim mCodBodega As String
      'Dim mCodUsuarioRet As String
      'Dim mCodUsuarioAut As String
      Dim mResultado As Boolean
   
   
      'TODO(Realizado): Reemplazar fValidaConfirmarConsumo x fValidaConfirmarIngreso. Redefinir las validaciones de esta funcion fValidaConfirmarIngreso
      'TODO(Realizado): Reemplazar los vEgresosCodProducto, vEgresosCodUbicacion, vEgresosCantidad
      If fValidaConfirmarIngreso() Then
         cantIngresos = FlexIngreso.Rows - 2
      
         ReDim vIngresosCodProducto(0 To cantIngresos - 1) As String
         ReDim vIngresosCodUbicacion(0 To cantIngresos - 1) As String
         ReDim vIngresosCantidad(0 To cantIngresos - 1) As Double
         
         
         For mJ = 2 To FlexIngreso.Rows - 1
            vIngresosCodProducto(mJ - 2) = FlexIngreso.TextMatrix(mJ, 6)
            vIngresosCodUbicacion(mJ - 2) = FlexIngreso.TextMatrix(mJ, 7)
            vIngresosCantidad(mJ - 2) = CDbl(Replace(FlexIngreso.TextMatrix(mJ, 3), ".", ","))
         Next
         
         mCodBodega = Left(Combo1.Text, 4)
   
         mResultado = True
         mObj.xInsIngreso vIngresosCodProducto(), vIngresosCodUbicacion(), vIngresosCantidad(), Trim(Text3.Text), mCodBodega, Trim(Right(MDI.mUser, 15)), mResultado
         
         If mResultado Then
            MsgBox "El Ingreso se ha realizado exitosamente.", vbInformation, "Ingresos"
            limpioFormulario
            actualizaFlagStockMinimo vIngresosCodProducto(), mCodBodega
         End If
      End If
   Else
      Unload Me
   End If
End Sub

Private Sub actualizaFlagStockMinimo(ByRef pvEgresosCodProducto() As String, ByVal pCodBodega As String)
   Dim mi As Integer

   For mi = LBound(pvEgresosCodProducto) To UBound(pvEgresosCodProducto)
      controlaFlagStockMinimo pvEgresosCodProducto(mi), pCodBodega
   Next
End Sub

Private Sub controlaFlagStockMinimo(ByVal pCodProducto As String, pCodBodega As String)

   Set mRec = mObj.oEjecutarSelect(" SELECT  M.CodProducto,  P.Descripcion AS Producto, U.CodBodega,  B.Descripcion AS Bodega,  SUM(Stock) AS Stock, " & _
      " IFNULL(SM.Stock_Min, 0) As Stock_Min,  SUM(Stock) - IFNULL(SM.Stock_Min, 0) AS StockMenosStockMin,  Med.Descripcion AS UnidadMedida, IFNULL(N.CodProducto,'XXXXXX') As ProductoNotificado " & _
      " FROM  " & _
      " Movimientos2 M " & _
      "  INNER JOIN " & _
      " Producto P ON M.CodProducto = P.Codigo " & _
      "  INNER JOIN " & _
      " Ubicaciones U ON  M.CodUbicacion = U.Codigo " & _
      "  INNER JOIN " & _
      " Bodegas B ON B.Codigo = U.CodBodega  " & _
      "  INNER JOIN " & _
      " UnidadMedida Med ON P.CodUnidadMedida = Med.Codigo " & _
      "  LEFT JOIN " & _
      " StocksMinimos SM ON SM.CodBodega = B.Codigo AND SM.CodProducto = M.CodProducto " & _
      "  LEFT JOIN  " & _
      "  StockMinimo_Notificaciones N ON N.CodProducto = M.CodProducto AND N.CodBodega = B.Codigo " & _
      " WHERE Fecha = (SELECT MAX(Fecha) " & _
      "                 From Movimientos2 " & _
      "                 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion) " & _
      " AND M.CodProducto = '" & pCodProducto & "' " & _
      " AND U.CodBodega = '" & pCodBodega & "' " & _
      " GROUP BY   M.CodProducto, P.Descripcion,U.CodBodega, B.Descripcion,Med.Descripcion, N.CodProducto;")

   'Si estoy por arriba del stock minimo y se notifico anteriormente entonces elimino flag de notificado.
   If CDbl(Replace(mRec!StockMenosStockMin, ".", ",")) > 0 And mRec!ProductoNotificado <> "XXXXXX" Then
         mObj.xDelStockMinimo_Notificaciones mRec!CodProducto, mRec!CodBodega
   End If
   mRec.Close
End Sub


Private Sub limpioFormulario()
   Dim mi As Integer

   Text1.Text = ""
   Text2.Text = ""
   Text3.Text = ""
  
   'Elimino los registros (de la consulta anterior) de la grilla superior
   For mi = FlexProduct.Rows To 3 Step -1
      FlexProduct.RemoveItem mi
    Next
   mRenglonProducto = 0
   
   'Elimino los registros de la grilla inferior
   For mi = FlexIngreso.Rows To 3 Step -1
      FlexIngreso.RemoveItem mi
   Next
   mRenglonIngreso = 0

End Sub

Private Sub FlexProduct_Click()
   Dim mi As Integer
   
   If FlexProduct.MouseRow > 0 Then
   
      If mRenglonProducto <> 0 Then
         FlexProduct.Row = mRenglonProducto
         For mi = 1 To FlexProduct.Cols - 1
            FlexProduct.Col = mi
            FlexProduct.CellBackColor = vbWhite
         Next
      End If
      
      mRenglonProducto = FlexProduct.MouseRow
   
      FlexProduct.Row = mRenglonProducto
      For mi = 1 To FlexProduct.Cols - 1
         FlexProduct.Col = mi
         FlexProduct.CellBackColor = &H80000003
      Next
      
      If mRenglonProducto > 1 Then
          mCodProducto = FlexProduct.TextMatrix(mRenglonProducto, 4)
      End If
   Else
      FlexProduct.Row = mRenglonProducto
      For mi = 1 To FlexProduct.Cols - 1
         FlexProduct.Col = mi
         FlexProduct.CellBackColor = vbWhite
      Next
      mRenglonProducto = 0
   End If
End Sub

Private Sub Form_Load()
   
   Inven011_frm.Width = 21270
   Inven011_frm.Height = 13950
   
   sAlinearForm Me
   
   'TODO(Realizado): Debe traer solo las bodegas que puede administrar el usuario. Tabla Futura Tabla: Usuarios-Bodegas (o sera mejor hacerlo por Almacen)
   Set mRec = mObj.oTabla("Bodegas", " where Fecha_Baja IS NULL and Codigo IN (SELECT CodBodega FROM Usuario_AccesoBodega WHERE CodUsuario = '" & Trim(Right(MDI.mUser, 15)) & "') order by Codigo")
   
   Do While Not mRec.EOF
      Combo1.AddItem "" & mRec!Codigo & " " & mRec!descripcion & ""
      mRec.MoveNext
   Loop
   mRec.Close

   FlexProduct.ColWidth(0) = 200
   FlexProduct.ColWidth(1) = 10700
   FlexProduct.ColWidth(2) = 4500
   FlexProduct.ColWidth(3) = 1500
   FlexProduct.ColWidth(4) = 1900
   FlexProduct.ColWidth(5) = 1250
   FlexProduct.ColWidth(6) = 0
   FlexProduct.ColWidth(7) = 0
   
   FlexProduct.TextMatrix(0, 1) = "Producto"
   FlexProduct.TextMatrix(0, 2) = "Ubicación"
   FlexProduct.TextMatrix(0, 3) = "Stock"
   FlexProduct.TextMatrix(0, 4) = "Unid.Medida"
   FlexProduct.TextMatrix(0, 5) = "Cód.Sap"
   FlexProduct.TextMatrix(0, 6) = "Cód. Producto"
   FlexProduct.TextMatrix(0, 7) = "Cód. Ubicacion"
   
   FlexProduct.RowHeight(1) = 0

   FlexIngreso.ColWidth(0) = 200
   FlexIngreso.ColWidth(1) = 10700
   FlexIngreso.ColWidth(2) = 4500
   FlexIngreso.ColWidth(3) = 1500
   FlexIngreso.ColWidth(4) = 1900
   FlexIngreso.ColWidth(5) = 1250
   FlexIngreso.ColWidth(6) = 0
   FlexIngreso.ColWidth(7) = 0
   
   
   FlexIngreso.TextMatrix(0, 1) = "Producto"
   FlexIngreso.TextMatrix(0, 2) = "Ubicación"
   FlexIngreso.TextMatrix(0, 3) = "Cantidad"
   FlexIngreso.TextMatrix(0, 4) = "Unid.Medida"
   FlexIngreso.TextMatrix(0, 5) = "Cód.Sap"
   FlexIngreso.TextMatrix(0, 6) = "Cód. Producto"
   FlexIngreso.TextMatrix(0, 7) = "Cód. Ubicacion"

   FlexIngreso.RowHeight(1) = 0

   cboListIndex = Combo1.ListIndex
End Sub

Private Function fValidaIngreso() As Boolean
   Dim mRet As Boolean
   Dim mMensajeError As String
   Dim mJ As Integer
   Dim mCantidaStock As Double
   Dim sStock As String
   Dim iStock As Double
   Dim mRec1 As New ADODB.Recordset
   Dim posInstr As Integer
   Dim qtyDecimales As Integer
    
   mRet = True
   
   
   If Trim(Text3.Text) = "" Then
      mRet = False
      mMensajeError = "Debe completar el Número de Orden de Compra"
   End If

   If mRet Then
      If Not IsNumeric(Trim(Text3.Text)) Then
         mRet = False
         mMensajeError = "La Orden de Compra debe ser un valor numérico !!"
      End If
   End If
      
   If mRet Then
      If Len(Trim(Text3.Text)) <> 10 Then
         mRet = False
           mMensajeError = "La Orden de Compra debe tener 10 caracteres numéricos !!"
      End If
   End If
   
   If mRet Then

      Set mRec1 = mObj.oEjecutarSelect("SELECT NroOC FROM Ingresos_H WHERE NroOC = " & Trim(Text3.Text) & "; ")
      If Not mRec1.EOF Then
         mRet = False
         mMensajeError = "Ya se han realizado ingresos para esa Orden de Compra !!!"
      End If
      mRec1.Close
   End If
   
   If mRet Then
      If mRenglonProducto = 0 Then
         mRet = False
         mMensajeError = "Debe seleccionar un Producto de la Grilla superior"
      End If
   End If
   
   If mRet Then
      If mRenglonProducto <> 0 And FlexProduct.TextMatrix(mRenglonProducto, 1) = "" Then
         mRet = False
         mMensajeError = "Debe seleccionar un Producto de la Grilla superior"
      End If
   End If
      
      
   If mRet Then
      If Trim(Text2.Text) = "" Then
         mRet = False
         mMensajeError = "Debe completar el campo: 'Cantidad'. "
      End If
   End If
      
   If mRet Then
      If Not IsNumeric(Replace(Text2.Text, ".", ",")) Then
         mRet = False
         mMensajeError = "La Cantidad ingresada no es un valor numérico"
      End If
   End If
   
   
   If mRet Then
      If CDbl(Replace(Trim(Text2.Text), ".", ",")) <= 0 Then
         mRet = False
         mMensajeError = "La Cantidad ingresada no puede ser menor o igual a cero."
      End If
   End If
   
   'Valido que no supere los 2 digitos decimales
   If mRet Then
      posInstr = InStr(1, Replace(Trim(Text2.Text), ".", ","), ",")
      
      If posInstr <> 0 Then
         qtyDecimales = Len(Right(Trim(Text2.Text), Len(Trim(Text2.Text)) - posInstr))
      End If
   
      If qtyDecimales > 2 Then
         mRet = False
         mMensajeError = "El campo 'Cantidad' solo admite hasta dos dígitos decimales."
      End If
   End If
   
   'Valido si el Producto y la Ubicacion seleccionados coinciden con un Egreso ya registrado
   If mRet Then
      For mJ = 2 To FlexIngreso.Rows - 1
         If FlexIngreso.TextMatrix(mJ, 6) = FlexProduct.TextMatrix(mRenglonProducto, 6) And FlexIngreso.TextMatrix(mJ, 7) = FlexProduct.TextMatrix(mRenglonProducto, 7) Then
            mRet = False
            mMensajeError = "El Producto y la Ubicación elegidos ya han sido seleccionados"
            mJ = 999
         End If
      Next
   End If
      
   
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   fValidaIngreso = mRet
End Function

Private Function fValidaConfirmarIngreso() As Boolean

  'Validado: 'Que se haya completado el campo Numero Vale.
  'Validado: 'Que el Numero de Vale sea un valor entero.
  'Validado: 'Que se haya completado el combo "Retirado por:"
  'Validado:  'Que esten chequeados alguna de las dos radio button (Option1 u Option2)
  'Validado: 'Que en la grilla inferior "Egresos" exista al menos un registro.
  'Validado: 'Que no exista el Registro (Nro Vale, TipoVale ) en la tabla Consumo_H
 
   Dim mRet As Boolean
   Dim mMensajeError As String
   Dim mRec1 As New ADODB.Recordset
   
   mRet = True
      
   If Trim(Text3.Text) = "" Then
      mRet = False
      mMensajeError = "Debe completar el Número de Orden de Compra"
   End If

   If mRet Then
      If Not IsNumeric(Trim(Text3.Text)) Then
         mRet = False
         mMensajeError = "La Orden de Compra debe ser un valor numérico !!"
      End If
   End If
      
   If mRet Then
      If Len(Trim(Text3.Text)) <> 10 Then
         mRet = False
           mMensajeError = "La Orden de Compra debe tener 10 caracteres numéricos !!"
      End If
   End If
            
   If mRet Then
      If FlexIngreso.Rows <= 2 Then
         mRet = False
         mMensajeError = "Al menos debe existir un registro en la Grilla Ingresos"
      End If
   End If
   
   If mRet Then

      Set mRec1 = mObj.oEjecutarSelect("SELECT NroOC FROM Ingresos_H WHERE NroOC = " & Trim(Text3.Text) & "; ")
      If Not mRec1.EOF Then
         mRet = False
         mMensajeError = "Ya se han realizado ingresos para esa Orden de Compra !!!"
      End If
      mRec1.Close
   End If

   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   
   fValidaConfirmarIngreso = mRet
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set mObj = Nothing
   Set mRec = Nothing
   ShowMenu 12, True, False
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
      If KeyAscii <> 46 Then
         KeyAscii = fNumeroKeyPress(KeyAscii)
      End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
         KeyAscii = fNumeroKeyPress(KeyAscii)
End Sub

Private Sub flexIngreso_Click()
   Dim mi As Integer
   
   If FlexIngreso.MouseRow > 0 Then
   
      If mRenglonIngreso <> 0 Then
         If FlexIngreso.Rows > mRenglonIngreso Then
            FlexIngreso.Row = mRenglonIngreso
            For mi = 1 To FlexIngreso.Cols - 1
               FlexIngreso.Col = mi
               FlexIngreso.CellBackColor = vbWhite
            Next
         End If
      End If
      
      mRenglonIngreso = FlexIngreso.MouseRow
   
      FlexIngreso.Row = mRenglonIngreso
      For mi = 1 To FlexIngreso.Cols - 1
         FlexIngreso.Col = mi
         FlexIngreso.CellBackColor = &H80000003
      Next
      
      If mRenglonIngreso > 1 Then
          mCodProducto = FlexIngreso.TextMatrix(mRenglonIngreso, 4)
      End If
   Else
      FlexIngreso.Row = mRenglonIngreso
      For mi = 1 To FlexProduct.Cols - 1
         FlexIngreso.Col = mi
         FlexIngreso.CellBackColor = vbWhite
      Next
      mRenglonIngreso = 0
   End If
End Sub

