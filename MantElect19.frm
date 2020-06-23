VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form MantElect19 
   Caption         =   "Ajustes de Inventario"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16965
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   16965
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   375
      Index           =   1
      Left            =   9480
      TabIndex        =   10
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Confirmar Ajuste"
      Height          =   375
      Index           =   0
      Left            =   5820
      TabIndex        =   9
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   0
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   1
      Left            =   3375
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6240
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ajuste"
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
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Top             =   6720
      Width           =   16695
      Begin MSFlexGridLib.MSFlexGrid FlexAjuste 
         Height          =   1095
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   16440
         _ExtentX        =   28998
         _ExtentY        =   1931
         _Version        =   327680
         Cols            =   9
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selección del Producto"
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
      Height          =   4575
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   16695
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   315
         Left            =   13080
         TabIndex        =   4
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   2520
         TabIndex        =   2
         Top             =   420
         Width           =   10455
      End
      Begin MSFlexGridLib.MSFlexGrid FlexProduct 
         Height          =   3375
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   16440
         _ExtentX        =   28998
         _ExtentY        =   5953
         _Version        =   327680
         Cols            =   8
      End
      Begin VB.Label Label3 
         Caption         =   "Contiene texto:"
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información del Ajuste"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16695
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   2520
         MaxLength       =   150
         TabIndex        =   3
         Top             =   840
         Width           =   13815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label5 
         Caption         =   "Motivo del ajuste:"
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Ajustar en Bodega:"
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   420
         Width           =   1455
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo de ajuste:"
      Height          =   255
      Left            =   4560
      TabIndex        =   19
      Top             =   6300
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Cantidad:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   6300
      Width           =   975
   End
End
Attribute VB_Name = "MantElect19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mObj As New clInven
Dim mRec As ADODB.Recordset
Dim mRenglonProducto As Integer
Dim mRenglonAjuste As Integer
Dim mCodProducto As String
Dim cboListIndex As Integer

Private Sub Combo1_Click()
   Dim mi As Integer
   If cboListIndex <> Combo1.ListIndex Then
      If (cboListIndex <> -1) Then
         If MsgBox("Si selecciona otra Bodega se perderán los datos cargados hasta el momento. ¿ Desea continuar ? ", vbYesNo, "Cambio de Bodega") = vbYes Then
            sMsgEspere Me, "Procesando datos...", True
            
            Text1.Text = ""
            Text2.Text = ""
            
            'Elimino los registros de la grilla superior (productos)
            For mi = FlexProduct.Rows To 3 Step -1
               FlexProduct.RemoveItem mi
            Next
            mRenglonProducto = 0
            
            'Elimino los registros de la grilla inferior (Ingresos)
            For mi = FlexAjuste.Rows To 3 Step -1
               FlexAjuste.RemoveItem mi
            Next
            mRenglonAjuste = 0
            sMsgEspere Me, "", False
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
   Dim mj As Integer
   
   sMsgEspere Me, "Buscando productos...", True
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
   
   sMsgEspere Me, "", False
End Sub

Private Sub Command2_Click(Index As Integer)
   Dim iStock As Double
   Dim mi As Integer
   Dim mRec1 As New ADODB.Recordset
   
   If Index = 0 Then
      If fValidaAjuste() Then
            FlexAjuste.AddItem vbTab & FlexProduct.TextMatrix(mRenglonProducto, 1) & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 2) & vbTab & Text2.Text & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 4) & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 5) & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 6) & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 7) & vbTab & Trim(Right(Combo2.Text, 1))
            Text2.Text = ""
            Text2.SetFocus
      End If
   Else
      If FlexAjuste.Rows > 2 And mRenglonAjuste > 1 Then
         FlexAjuste.RemoveItem (mRenglonAjuste)
      End If
      
      mRenglonAjuste = 0
   End If
End Sub

'Boton de "Confirmacion de Ajuste" o "Salir2
Private Sub Command3_Click(Index As Integer)
   If Index = 0 Then
      Dim vAjustesCodProducto() As String
      Dim vAjustesCodUbicacion() As String
      Dim vAjustesCantidad() As Double
      Dim vAjustesTipoAjuste() As String
      Dim cantAjustes As Integer
      Dim mj As Integer
      Dim mCodBodega As String
      Dim mResultado As Boolean
      Dim mMotivo As String
      Dim mProducto As String
   
      If fValidaConfirmarAjuste() Then
         cantAjustes = FlexAjuste.Rows - 2
      
         ReDim vAjustesCodProducto(0 To cantAjustes - 1) As String
         ReDim vAjustesCodUbicacion(0 To cantAjustes - 1) As String
         ReDim vAjustesCantidad(0 To cantAjustes - 1) As Double
         ReDim vAjustesTipoAjuste(0 To cantAjustes - 1) As String
         
         
         sMsgEspere Me, "Procesando datos...", True
         
         For mj = 2 To FlexAjuste.Rows - 1
            vAjustesCodProducto(mj - 2) = FlexAjuste.TextMatrix(mj, 6)
            vAjustesCodUbicacion(mj - 2) = FlexAjuste.TextMatrix(mj, 7)
            vAjustesCantidad(mj - 2) = CDbl(Replace(FlexAjuste.TextMatrix(mj, 3), ".", ","))
            vAjustesTipoAjuste(mj - 2) = FlexAjuste.TextMatrix(mj, 8)
         Next
         
         mCodBodega = Left(Combo1.Text, 4)
         mMotivo = Trim(Text3.Text)
         
         mResultado = True
         mObj.xInsAjuste vAjustesCodProducto(), vAjustesCodUbicacion(), vAjustesCantidad(), vAjustesTipoAjuste(), Trim(Text3.Text), Trim(Right(MDI.mUser, 15)), mResultado
         
         If mResultado Then
            limpioFormulario
            updFlagStockMinimo vAjustesCodProducto(), vAjustesTipoAjuste(), mCodBodega
            notificaAjustes vAjustesCodProducto(), vAjustesCodUbicacion(), vAjustesCantidad(), vAjustesTipoAjuste(), mCodBodega, mMotivo
            sMsgEspere Me, "", False
            MsgBox "El Ajuste se ha realizado exitosamente.", vbInformation, "Ajustes"
         End If
         sMsgEspere Me, "", False
      End If
   Else
      Unload Me
   End If
End Sub

Private Sub updFlagStockMinimo(ByRef pvAjustesCodProducto() As String, ByRef pvAjustesTipoAjuste() As String, ByVal pCodBodega As String)
   Dim mi As Integer
   For mi = LBound(pvAjustesCodProducto) To UBound(pvAjustesCodProducto)
      If pvAjustesTipoAjuste(mi) = "I" Then
         controlaFlagStockMinimo pvAjustesCodProducto(mi), pCodBodega
      Else
         VerificaStockMinYnotifica pvAjustesCodProducto(mi), pCodBodega
      End If
   Next
End Sub

Private Sub notificaAjustes(ByRef pvAjustesCodProducto() As String, ByRef pvAjustesCodUbicacion() As String, ByRef pvAjustesCantidad() As Double, ByRef pvAjustesTipoAjuste() As String, ByVal pCodBodega As String, pMotivo As String)
   Dim mi As Integer
   For mi = LBound(pvAjustesCodProducto) To UBound(pvAjustesCodProducto)
      notificaUnAjuste pvAjustesCodProducto(mi), pvAjustesCodUbicacion(mi), pvAjustesCantidad(mi), pvAjustesTipoAjuste(mi), pCodBodega, pMotivo
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

Private Sub VerificaStockMinYnotifica(ByVal pCodProducto As String, pCodBodega As String)

   Dim mRec1 As ADODB.Recordset
   Dim mListaDestinatarios As String
   Dim mTextoMail As String
   
   mListaDestinatarios = ""
   mTextoMail = ""


   Set mRec = mObj.oEjecutarSelect(" SELECT  M.CodProducto, P.CodigoSap, P.Descripcion AS Producto, U.CodBodega,  B.Descripcion AS Bodega,  SUM(Stock) AS Stock, " & _
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

      Set mRec1 = mObj.oEjecutarSelect(" SELECT DISTINCT P.Email FROM " & _
                                       "    Usuario_Bodega_Notificacion U " & _
                                       " INNER JOIN " & _
                                       "  Personal P ON P.CodUsuario = U.CodUsuario " & _
                                       " WHERE U.CodBodega = '& pCodBodega &'; ")
                                 
   Do While Not mRec1.EOF
      mListaDestinatarios = mListaDestinatarios & mRec!Email & ";"
      mRec1.MoveNext
   Loop
   mRec1.Close

   'Si estoy por debajo del stock minimo y no se notifico anteriormente entonces enivo correo y cargo flag de notificado.
   If CDbl(Replace(mRec!StockMenosStockMin, ".", ",")) <= 0 And mRec!ProductoNotificado = "XXXXXX" Then
      Set mRec1 = mObj.oEjecutarSelect(" SELECT DISTINCT P.Email FROM " & _
                                    "    Usuario_Bodega_Notificacion U " & _
                                    " INNER JOIN " & _
                                    "  Personal P ON P.CodUsuario = U.CodUsuario " & _
                                    " WHERE U.CodBodega = '" & pCodBodega & "'; ")
                                    
      Do While Not mRec1.EOF
         mListaDestinatarios = mListaDestinatarios & mRec1!Email & ";"
         mRec1.MoveNext
      Loop
      mRec1.Close
      
      mTextoMail = vbCrLf & _
                  " A continuación se detallan los datos del producto que ha llegado a su Stock Mínimo: " & vbCrLf & _
                   vbCrLf & _
                   vbCrLf & _
                  Space(5) & "Cód. Producto: " & mRec!CodProducto & vbCrLf & _
                  Space(5) & "Código SAP: " & mRec!CodigoSap & vbCrLf & _
                  Space(5) & "Producto: " & mRec!Producto & vbCrLf & _
                  Space(5) & "Bodega: " & mRec!Bodega & vbCrLf & _
                  Space(5) & "Stock Actual: " & Format(mRec!Stock, "#.00") & " " & mRec!UnidadMedida & vbCrLf & _
                  Space(5) & "Stock Mínimo: " & Format(mRec!Stock_Min, "#.00") & " " & mRec!UnidadMedida & vbCrLf
                  
      
      If fEnviar_Mail_CDO("", mListaDestinatarios, "ausolmail@ausol.com.ar", " Sist. Global - Inventario: Stock mínimo alcanzó su límite", mTextoMail, "", "587", "ausolmail@ausol.com.ar", "sgedosmildiecisiete1$", True, False) Then
         mObj.xInsStockMinimo_Notificaciones mRec!CodProducto, mRec!CodBodega
      End If
   End If
   mRec.Close
End Sub


Private Sub notificaUnAjuste(ByVal pCodProducto As String, pCodUbicacion As String, pCodCantidad As Double, pTipoAjuste As String, pCodBodega As String, pMotivoAjuste As String)
   Dim mRec1 As ADODB.Recordset
   Dim mListaDestinatarios As String
   Dim mTextoMail As String
   Dim bResult As Boolean
   Dim ProductoDescr As String
   Dim CodSap As String
   Dim BodegaDescr   As String
   Dim UbicacionDescr As String
   Dim TipoAjusteDescr As String
   Dim PersonaApe As String
   Dim PersonaNombre As String
   
   mListaDestinatarios = ""
   mTextoMail = ""
   
   Set mRec1 = mObj.oEjecutarSelect(" SELECT DISTINCT P.Email FROM " & _
                                    "    Usuario_Bodega_Notificacion U " & _
                                    " INNER JOIN " & _
                                    "  Personal P ON P.CodUsuario = U.CodUsuario " & _
                                    " WHERE U.CodBodega = '" & pCodBodega & "'; ")
                                    
   ProductoDescr = mObj.sTablaDescr("Producto", "Codigo ='" & pCodProducto & "'", 1)
   CodSap = mObj.sTablaDescr("Producto", "Codigo ='" & pCodProducto & "'", 2)
   BodegaDescr = mObj.sTablaDescr("Bodegas", "Codigo ='" & pCodBodega & "'", 1)
   UbicacionDescr = mObj.sTablaDescr("Ubicaciones", "Codigo ='" & pCodUbicacion & "'", 1)
   If pTipoAjuste = "I" Then
      TipoAjusteDescr = "Ingreso"
   Else
      TipoAjusteDescr = "Egreso"
   End If
   
   PersonaApe = mObj.sTablaDescr("Personal", "CodUsuario ='" & Trim(Right(MDI.mUser, 15)) & "'", 1)
   PersonaNombre = mObj.sTablaDescr("Personal", "CodUsuario ='" & Trim(Right(MDI.mUser, 15)) & "'", 2)
   
   Do While Not mRec1.EOF
      mListaDestinatarios = mListaDestinatarios & mRec1!Email & ";"
      mRec1.MoveNext
   Loop
   mRec1.Close
   
   mTextoMail = vbCrLf & _
      " A continuación se detallan los datos del ajuste realizado: " & vbCrLf & _
      vbCrLf & _
      vbCrLf & _
      Space(5) & "Código Producto: " & pCodProducto & vbCrLf & _
      Space(5) & "Código Sap: " & CodSap & vbCrLf & _
      Space(5) & "Nombre Producto: " & ProductoDescr & vbCrLf & _
      Space(5) & "Bodega: " & pCodBodega & vbCrLf & _
      Space(5) & "Nombre Bodega: " & BodegaDescr & vbCrLf & _
      Space(5) & "Ubicación: " & pCodUbicacion & vbCrLf & _
      Space(5) & "Nombre Ubicación: " & UbicacionDescr & vbCrLf & _
      Space(5) & "Tipo Ajuste: " & TipoAjusteDescr & vbCrLf & _
      Space(5) & "Cantidad: " & pCodCantidad & vbCrLf & _
      Space(5) & "Motivo Ajuste: " & pMotivoAjuste & vbCrLf & _
      Space(5) & "Ajuste realizado por: " & PersonaNombre & " " & PersonaApe
                  
      bResult = fEnviar_Mail_CDO("", mListaDestinatarios, "ausolmail@ausol.com.ar", " Sist. Global - Inventario: Notificación de Ajuste", mTextoMail, "", "587", "ausolmail@ausol.com.ar", "sgedosmildiecisiete1$", True, False)
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
   For mi = FlexAjuste.Rows To 3 Step -1
      FlexAjuste.RemoveItem mi
   Next
   mRenglonAjuste = 0
   
   Combo2.Clear
   Combo2.AddItem "Positivo" & Space(50) & "I"
   Combo2.AddItem "Negativo" & Space(50) & "E"
   

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
   
   Me.Width = 17085
   Me.Height = 9750
   
   sAlinearForm Me
   
   Set mRec = mObj.oTabla("Bodegas", " where Fecha_Baja IS NULL and Codigo IN (SELECT CodBodega FROM Usuario_AccesoBodega WHERE CodUsuario = '" & Trim(Right(MDI.mUser, 15)) & "') order by Codigo")
   
   Do While Not mRec.EOF
      Combo1.AddItem "" & mRec!Codigo & " " & mRec!descripcion & ""
      mRec.MoveNext
   Loop
   mRec.Close


   Combo2.AddItem "Positivo" & Space(50) & "I"
   Combo2.AddItem "Negativo" & Space(50) & "E"

   FlexProduct.ColWidth(0) = 200
   FlexProduct.ColWidth(1) = 9500
   FlexProduct.ColWidth(2) = 2100
   FlexProduct.ColWidth(3) = 1200
   FlexProduct.ColWidth(4) = 1900
   FlexProduct.ColWidth(5) = 0
   FlexProduct.ColWidth(6) = 1150
   FlexProduct.ColWidth(7) = 0
   
   FlexProduct.TextMatrix(0, 1) = "Producto"
   FlexProduct.TextMatrix(0, 2) = "Ubicación"
   FlexProduct.TextMatrix(0, 3) = "Stock"
   FlexProduct.TextMatrix(0, 4) = "Unid.Medida"
   FlexProduct.TextMatrix(0, 5) = "Cód.Sap"
   FlexProduct.TextMatrix(0, 6) = "Cód. Producto"
   FlexProduct.TextMatrix(0, 7) = "Cód. Ubicacion"
   
   FlexProduct.RowHeight(1) = 0

   FlexAjuste.ColWidth(0) = 200
   FlexAjuste.ColWidth(1) = 9500
   FlexAjuste.ColWidth(2) = 2100
   FlexAjuste.ColWidth(3) = 1200
   FlexAjuste.ColWidth(4) = 1900
   FlexAjuste.ColWidth(5) = 0
   FlexAjuste.ColWidth(6) = 1150
   FlexAjuste.ColWidth(7) = 0
   FlexAjuste.ColWidth(8) = 0
   
   FlexAjuste.TextMatrix(0, 1) = "Producto"
   FlexAjuste.TextMatrix(0, 2) = "Ubicación"
   FlexAjuste.TextMatrix(0, 3) = "Cantidad"
   FlexAjuste.TextMatrix(0, 4) = "Unid.Medida"
   FlexAjuste.TextMatrix(0, 5) = "Cód.Sap"
   FlexAjuste.TextMatrix(0, 6) = "Cód. Producto"
   FlexAjuste.TextMatrix(0, 7) = "Cód. Ubicacion"
   FlexAjuste.TextMatrix(0, 8) = "Cód. TipoAjuste"

   FlexAjuste.RowHeight(1) = 0

   cboListIndex = Combo1.ListIndex
End Sub

Private Function fValidaAjuste() As Boolean
   Dim mRet As Boolean
   Dim mMensajeError As String
   Dim mj As Integer
   Dim mCantidaStock As Double
   Dim sStock As String
   Dim iStock As Double
   Dim mRec1 As New ADODB.Recordset
   Dim posInstr As Integer
   Dim qtyDecimales As Integer
    
   mRet = True
   
   If mRenglonProducto = 0 Then
      mRet = False
      mMensajeError = "Debe seleccionar un Producto de la Grilla superior"
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
   
   If mRet Then
      If Combo2.ListIndex = -1 Then
         
         mRet = False
         mMensajeError = "Debe seleccionar el tipo de ajuste."
      
      End If
   End If

   'Valido que solo exista un solo registro en la grilla inferior
   If mRet Then
      If FlexAjuste.Rows = 3 Then
         mRet = False
         mMensajeError = "No se puede realizar mas de un ajuste por operación."
      End If
   End If
   
   'Si el tipo de ajuste es negativo entonces valido si el saldo del stock es insuficiente para ese Producto/Ubicación
   If mRet Then
      If Combo2.ListIndex <> -1 And Trim(Right(Combo2.Text, 1)) = "E" Then
            Set mRec1 = mObj.oEjecutarSelect("SELECT Fecha,CodProducto,CodUbicacion, IFNULL(Stock,0) As Stock " & _
                                                      " FROM Movimientos2 M " & _
                                                      " WHERE CodProducto  = '" & FlexProduct.TextMatrix(mRenglonProducto, 6) & "' and CodUbicacion = '" & FlexProduct.TextMatrix(mRenglonProducto, 7) & "'" & _
                                                      " AND Fecha = (SELECT Max(Fecha) FROM Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion)")
            If Not mRec1.EOF Then
               iStock = mRec1!Stock
            Else
               iStock = 0
            End If
            mRec1.Close
            
            If CDbl(Replace(Trim(Text2.Text), ".", ",")) > iStock Then
               mRet = False
               mMensajeError = "El stock es insuficiente para ese Producto en esa Ubicación"
            End If
      End If
   End If
      
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   fValidaAjuste = mRet
End Function

Private Function fValidaConfirmarAjuste() As Boolean

   Dim mRet As Boolean
   Dim mMensajeError As String
   
   mRet = True
      
   If Trim(Text3.Text) = "" Then
      mRet = False
      mMensajeError = "Debe completar el Motivo del Ajuste"
   End If

            
   If mRet Then
      If FlexAjuste.Rows <= 2 Then
         mRet = False
         mMensajeError = "Al menos debe existir un registro en la Grilla Ajustes"
      End If
   End If
   
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   
   fValidaConfirmarAjuste = mRet
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

Private Sub FlexAjuste_Click()
   Dim mi As Integer
   
   If FlexAjuste.MouseRow > 0 Then
   
      If mRenglonAjuste <> 0 Then
         If FlexAjuste.Rows > mRenglonAjuste Then
            FlexAjuste.Row = mRenglonAjuste
            For mi = 1 To FlexAjuste.Cols - 1
               FlexAjuste.Col = mi
               FlexAjuste.CellBackColor = vbWhite
            Next
         End If
      End If
      
      mRenglonAjuste = FlexAjuste.MouseRow
   
      FlexAjuste.Row = mRenglonAjuste
      For mi = 1 To FlexAjuste.Cols - 1
         FlexAjuste.Col = mi
         FlexAjuste.CellBackColor = &H80000003
      Next
      
      If mRenglonAjuste > 1 Then
          mCodProducto = FlexAjuste.TextMatrix(mRenglonAjuste, 4)
      End If
   Else
      FlexAjuste.Row = mRenglonAjuste
      For mi = 1 To FlexProduct.Cols - 1
         FlexAjuste.Col = mi
         FlexAjuste.CellBackColor = vbWhite
      Next
      mRenglonAjuste = 0
   End If
End Sub
