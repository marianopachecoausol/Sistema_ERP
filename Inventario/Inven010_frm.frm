VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Inven010_frm 
   Caption         =   "Consumo de Materiales"
   ClientHeight    =   9240
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   16965
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   16965
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   375
      Index           =   1
      Left            =   9450
      TabIndex        =   23
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   1
      Left            =   3375
      Picture         =   "Inven010_frm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Index           =   0
      Left            =   2880
      Picture         =   "Inven010_frm.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Confirmar Consumo"
      Height          =   375
      Index           =   0
      Left            =   5820
      TabIndex        =   9
      Top             =   8640
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1320
      TabIndex        =   7
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Caption         =   "Egresos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3015
      Left            =   120
      TabIndex        =   15
      Top             =   5400
      Width           =   16695
      Begin MSFlexGridLib.MSFlexGrid FlexEgreso 
         Height          =   2535
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   16440
         _ExtentX        =   28998
         _ExtentY        =   4471
         _Version        =   327680
         Cols            =   8
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   3375
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   16695
      Begin MSFlexGridLib.MSFlexGrid FlexProduct 
         Height          =   2295
         Left            =   120
         TabIndex        =   14
         Top             =   945
         Width           =   16440
         _ExtentX        =   28998
         _ExtentY        =   4048
         _Version        =   327680
         Cols            =   8
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   315
         Left            =   12360
         TabIndex        =   6
         Top             =   400
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   400
         Width           =   10455
      End
      Begin VB.Label Label3 
         Caption         =   "Contiene texto:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información del Consumo"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16695
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   720
         Width           =   3375
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   12720
         MaxLength       =   9
         TabIndex        =   2
         Top             =   440
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Vale de retiro múltiple"
         Height          =   375
         Left            =   14160
         TabIndex        =   4
         Top             =   620
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Vale a cargo/recambio"
         Height          =   375
         Left            =   14160
         TabIndex        =   3
         Top             =   200
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Inven010_frm.frx":0614
         Left            =   1800
         List            =   "Inven010_frm.frx":0616
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   440
         Width           =   3735
      End
      Begin VB.Label Label6 
         Caption         =   "Autorizado por:"
         Height          =   255
         Left            =   6240
         TabIndex        =   22
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Vale número:"
         Height          =   255
         Left            =   11400
         TabIndex        =   19
         Top             =   495
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Retirado por:"
         Height          =   255
         Left            =   6240
         TabIndex        =   16
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Retirar de Bodega:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   500
         Width           =   1575
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Cantidad:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4980
      Width           =   975
   End
End
Attribute VB_Name = "Inven010_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mObj As New clInven
Dim mRec As ADODB.Recordset
Dim mRenglonProducto As Integer
Dim mRenglonEgreso As Integer
Dim mCodProducto As String
Dim cboListIndex As Integer

Private Sub Combo1_Click_OLD()
   Dim mi As Integer
   
   Combo2.Enabled = True
   Combo3.Enabled = True
   
   If cboListIndex <> Combo1.ListIndex Then
      sLlenoUsuariosRet
      sLlenoUsuariosAut
      If (cboListIndex <> -1) Then
         
         If MsgBox("Si selecciona otra Bodega se perderán los consumos cargados hasta el momento. ¿ Desea continuar ? ", vbYesNo, "Cambio de Bodega") = vbYes Then
            Text1.Text = ""
            Text2.Text = ""
            
            'Elimino los registros de la grilla superior (productos)
            For mi = FlexProduct.Rows To 3 Step -1
               FlexProduct.RemoveItem mi
            Next
            
            'Elimino los registros de la grilla inferior (consumos)
            For mi = FlexEgreso.Rows To 3 Step -1
               FlexEgreso.RemoveItem mi
            Next
            
            mRenglonProducto = 0
            mRenglonEgreso = 0
         Else
            Combo1.ListIndex = cboListIndex
            sLlenoUsuariosRet
            sLlenoUsuariosAut
         End If
         
         cboListIndex = Combo1.ListIndex
      
      Else
         cboListIndex = Combo1.ListIndex
      End If
   
   End If
End Sub

Private Sub Combo1_Click()
   Dim mi As Integer
   
   Combo2.Enabled = True
   Combo3.Enabled = True
   
   If cboListIndex <> Combo1.ListIndex Then
      sLlenoUsuariosRet
      sLlenoUsuariosAut
      If (cboListIndex <> -1) Then
         'Si tengo algun registro en la grilla inferior(Egresos)
         If FlexEgreso.Rows > 2 Then
            If MsgBox("Si selecciona otra Bodega se perderán los consumos cargados hasta el momento. ¿ Desea continuar ? ", vbYesNo, "Cambio de Bodega") = vbYes Then
                
               sMsgEspere Me, "Procesando datos...", True
               Text1.Text = ""
               Text2.Text = ""
               
               'Elimino los registros de la grilla superior (productos)
               For mi = FlexProduct.Rows To 3 Step -1
                  FlexProduct.RemoveItem mi
               Next
               
               'Elimino los registros de la grilla inferior (consumos)
               For mi = FlexEgreso.Rows To 3 Step -1
                  FlexEgreso.RemoveItem mi
               Next
               
               mRenglonProducto = 0
               mRenglonEgreso = 0
               sMsgEspere Me, "", False
               
            Else
               Combo1.ListIndex = cboListIndex
               sLlenoUsuariosRet
               sLlenoUsuariosAut
            End If
         Else
            Text1.Text = ""
            Text2.Text = ""
               
            'Elimino los registros de la grilla superior (productos)
            For mi = FlexProduct.Rows To 3 Step -1
               FlexProduct.RemoveItem mi
            Next
            
         End If
         
         cboListIndex = Combo1.ListIndex
      
      Else
         cboListIndex = Combo1.ListIndex
      End If
   
   End If
End Sub

Private Sub sLlenoUsuariosRet()
Dim mCodBodega As String
Dim mObj2 As New clInven
Dim mRec2 As New ADODB.Recordset
   
   mCodBodega = Trim(Left(Combo1.Text, 4))
   Combo2.Clear
   
   Set mRec2 = mObj.oEjecutarSelect("SELECT CONCAT (P.Apellido,',', P.Nombres) AS Descripcion,P.CodUsuario AS CodUsuario FROM " & _
   " UsuariosRet_Bodegas UB " & _
   " Inner Join " & _
   " Personal P ON UB.CodUsuario = P.CodUsuario " & _
   " WHERE UB.CodBodega = '" & mCodBodega & "' AND P.Fecha_Baja IS NULL " & _
   " ORDER BY P.Apellido;")
   
   
   Do While Not mRec2.EOF
      Combo2.AddItem mRec2!descripcion & Space(60) & mRec2!CodUsuario
      mRec2.MoveNext
   Loop
   mRec2.Close
   Set mObj2 = Nothing
   Set mRec2 = Nothing
End Sub


Private Sub sLlenoUsuariosAut()
Dim mCodBodega As String
Dim mObj2 As New clInven
Dim mRec2 As New ADODB.Recordset
   
   mCodBodega = Trim(Left(Combo1.Text, 4))
   Combo3.Clear
   
   Set mRec2 = mObj.oEjecutarSelect("SELECT CONCAT (P.Apellido,',', P.Nombres) AS Descripcion,P.CodUsuario AS CodUsuario FROM " & _
   " UsuariosAut_Bodegas UB " & _
   " Inner Join " & _
   " Personal P ON UB.CodUsuario = P.CodUsuario " & _
   " WHERE UB.CodBodega = '" & mCodBodega & "' AND P.Fecha_Baja IS NULL " & _
   " ORDER BY P.Apellido;")
   
   
   Do While Not mRec2.EOF
      Combo3.AddItem mRec2!descripcion & Space(60) & mRec2!CodUsuario
      mRec2.MoveNext
   Loop
   mRec2.Close
   Set mObj2 = Nothing
   Set mRec2 = Nothing
End Sub

Private Sub Command1_Click()
   Dim mi As Integer
   Dim mj As Integer
   
   mRenglonProducto = 0
   
   sMsgEspere Me, "Buscando productos...", True
   
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
     
   'Si presiono el boton Buscar y algun "producto/ubicacion" de la grilla de arriba, esta en la grilla inferior
   'entonces que me actualice en la grilla superior el stock de ese o esos "producto/ubicacion" descontando el consumo de la grilla inferior
   For mi = 2 To FlexProduct.Rows - 1
      For mj = 2 To FlexEgreso.Rows - 1
         If FlexProduct.TextMatrix(mi, 6) = FlexEgreso.TextMatrix(mj, 6) And FlexProduct.TextMatrix(mi, 7) = FlexEgreso.TextMatrix(mj, 7) Then
            FlexProduct.TextMatrix(mi, 3) = CDbl(Replace(Trim(FlexProduct.TextMatrix(mi, 3)), ".", ",")) - CDbl(Replace(Trim(FlexEgreso.TextMatrix(mj, 3)), ".", ","))
            mj = 999
         End If
      Next
   Next
   
   sMsgEspere Me, "", False
   
End Sub

Private Sub Command2_Click(Index As Integer)
   Dim iStock As Double
   Dim mi As Integer
   Dim mRec1 As New ADODB.Recordset
   
   If Index = 0 Then
      If fValidaEgreso() Then
            FlexEgreso.AddItem vbTab & FlexProduct.TextMatrix(mRenglonProducto, 1) & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 2) & vbTab & Text2.Text & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 4) & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 5) & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 6) & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 7)
            FlexProduct.TextMatrix(mRenglonProducto, 3) = CDbl(Replace(Trim(FlexProduct.TextMatrix(mRenglonProducto, 3)), ".", ",")) - CDbl(Replace(Trim(Text2.Text), ".", ","))
            Text2.Text = ""
            Text2.SetFocus
      End If
   Else
      For mi = 2 To FlexProduct.Rows - 1
      
         If FlexProduct.TextMatrix(mi, 6) = FlexEgreso.TextMatrix(mRenglonEgreso, 6) And FlexProduct.TextMatrix(mi, 7) = FlexEgreso.TextMatrix(mRenglonEgreso, 7) Then
            Set mRec1 = mObj.oEjecutarSelect("SELECT Fecha,CodProducto,CodUbicacion, IFNULL(Stock,0) As Stock " & _
                                       " FROM Movimientos2 M " & _
                                       " WHERE CodProducto  = '" & FlexEgreso.TextMatrix(mRenglonEgreso, 6) & "' and CodUbicacion = '" & FlexEgreso.TextMatrix(mRenglonEgreso, 7) & "'" & _
                                       " AND Fecha = (SELECT Max(Fecha) FROM Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion)")
      
            If Not mRec1.EOF Then
               iStock = mRec1!Stock
            Else
               iStock = 0
            End If
            mRec1.Close
            
            FlexProduct.TextMatrix(mi, 3) = iStock
           
            mi = 9999
         End If
      Next
      
      If FlexEgreso.Rows > 2 And mRenglonEgreso > 1 Then
         FlexEgreso.RemoveItem (mRenglonEgreso)
      End If
      
      mRenglonEgreso = 0
   End If
End Sub

'Boton de confirmacion de Consumo de materiales
Private Sub Command3_Click(Index As Integer)
   If Index = 0 Then
      Dim vEgresosCodProducto() As String
      Dim vEgresosCodUbicacion() As String
      Dim vEgresosCantidad() As Double
      Dim cantEgresos As Integer
      Dim mj As Integer
      Dim mCodTipoVale As String
      Dim mCodBodega As String
      Dim mCodUsuarioRet As String
      Dim mCodUsuarioAut As String
      Dim mResultado As Boolean
   
      If fValidaConfirmarConsumo() Then
         cantEgresos = FlexEgreso.Rows - 2
      
         ReDim vEgresosCodProducto(0 To cantEgresos - 1) As String
         ReDim vEgresosCodUbicacion(0 To cantEgresos - 1) As String
         ReDim vEgresosCantidad(0 To cantEgresos - 1) As Double
         
         sMsgEspere Me, "Procesando datos...", True
         
         For mj = 2 To FlexEgreso.Rows - 1
            vEgresosCodProducto(mj - 2) = FlexEgreso.TextMatrix(mj, 6)
            vEgresosCodUbicacion(mj - 2) = FlexEgreso.TextMatrix(mj, 7)
            vEgresosCantidad(mj - 2) = CDbl(Replace(FlexEgreso.TextMatrix(mj, 3), ".", ","))
         Next
         
         If Option1.Value Then
            mCodTipoVale = "C"
         Else
            mCodTipoVale = "M"
         End If
         
         mCodBodega = Left(Combo1.Text, 4)
         mCodUsuarioRet = Trim(Right(Combo2.Text, 25))
         mCodUsuarioAut = Trim(Right(Combo3.Text, 25))
         mResultado = True
         'OK 'Inserto en Consumo_H ->OK: FALTA TIPOVALE,CODBODETA,USUARIORETIRA,USURIOSIST
         mObj.xInsEgreso vEgresosCodProducto(), vEgresosCodUbicacion(), vEgresosCantidad(), Trim(Text3.Text), mCodTipoVale, mCodBodega, mCodUsuarioRet, mCodUsuarioAut, Trim(Right(MDI.mUser, 15)), mResultado
         'OK 'Inserto en Consumo_Det
         
         If mResultado Then
             sMsgEspere Me, "", False
            limpioFormulario
            VerificaStockMin vEgresosCodProducto(), mCodBodega
            MsgBox "El consumo se ha realizado exitosamente", vbInformation, "Consumo"
         End If
         sMsgEspere Me, "", False
         
         'Validado: 'Que se haya completado el campo Numero Vale.
         'Validado: 'Que el Numero de Vale sea un valor entero.
         'Validado: 'Que se haya completado el combo "Retirado por:"
         'Validado: 'Que esten chequeados alguna de las dos radio button (Option1 u Option2)
         'Validado: 'Que en la grilla inferior "Egresos" exista al menos un registro.
         'Validado: 'Que no exista el Registro (Nro Vale, TipoVale ) en la tabla Consumo_H
      End If
   Else
      Unload Me
   End If
   
End Sub


Private Sub VerificaStockMin(ByRef pvEgresosCodProducto() As String, ByVal pCodBodega As String)
   Dim mi As Integer

   For mi = LBound(pvEgresosCodProducto) To UBound(pvEgresosCodProducto)
      VerificaStockMinYnotifica pvEgresosCodProducto(mi), pCodBodega
   Next
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


Private Sub limpioFormulario()
   Dim mi As Integer

   Text1.Text = ""
   Text2.Text = ""
   Text3.Text = ""
   
   Option1.Value = False
   Option2.Value = False
   
   'Elimino los registros (de la consulta anterior) de la grilla superior
   For mi = FlexProduct.Rows To 3 Step -1
      FlexProduct.RemoveItem mi
   Next
     
   mRenglonProducto = 0

   'Elimino los registros de la grilla inferior
   For mi = FlexEgreso.Rows To 3 Step -1
      FlexEgreso.RemoveItem mi
   Next
     
   mRenglonEgreso = 0
   
   Combo1.Clear
   Set mRec = mObj.oTabla("Bodegas", " where Fecha_Baja IS NULL and Codigo IN (SELECT CodBodega FROM Usuario_AccesoBodega WHERE CodUsuario = '" & Trim(Right(MDI.mUser, 15)) & "') order by Codigo")
   
   
   Do While Not mRec.EOF
      Combo1.AddItem "" & mRec!Codigo & " " & mRec!descripcion & ""
      mRec.MoveNext
   Loop
   mRec.Close
   cboListIndex = Combo1.ListIndex
   
   Combo2.Clear
   Combo3.Clear
   Combo2.Enabled = False
   Combo3.Enabled = False
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
   Inven010_frm.Width = 17085
   Inven010_frm.Height = 9750
   
   sAlinearForm Me
   
   Combo2.Enabled = False
   Combo3.Enabled = False
   
   Set mRec = mObj.oTabla("Bodegas", " where Fecha_Baja IS NULL and Codigo IN (SELECT CodBodega FROM Usuario_AccesoBodega WHERE CodUsuario = '" & Trim(Right(MDI.mUser, 15)) & "') order by Codigo")

   Do While Not mRec.EOF
      Combo1.AddItem "" & mRec!Codigo & " " & mRec!descripcion & ""
      mRec.MoveNext
   Loop
   mRec.Close

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

   FlexEgreso.ColWidth(0) = 200
   FlexEgreso.ColWidth(1) = 9500
   FlexEgreso.ColWidth(2) = 2100
   FlexEgreso.ColWidth(3) = 1200
   FlexEgreso.ColWidth(4) = 1900
   FlexEgreso.ColWidth(5) = 0
   FlexEgreso.ColWidth(6) = 1150
   FlexEgreso.ColWidth(7) = 0
   
   FlexEgreso.TextMatrix(0, 1) = "Producto"
   FlexEgreso.TextMatrix(0, 2) = "Ubicación"
   FlexEgreso.TextMatrix(0, 3) = "Cantidad"
   FlexEgreso.TextMatrix(0, 4) = "Unid.Medida"
   FlexEgreso.TextMatrix(0, 5) = "Cód.Sap"
   FlexEgreso.TextMatrix(0, 6) = "Cód. Producto"
   FlexEgreso.TextMatrix(0, 7) = "Cód. Ubicacion"

   FlexEgreso.RowHeight(1) = 0

   cboListIndex = Combo1.ListIndex
End Sub

Private Function fValidaEgreso() As Boolean
   Dim mRet As Boolean
   Dim mMensajeError As String
   Dim mj As Integer
   Dim mCantidaStock As Double
   Dim sStock As String
   Dim iStock As Double
   Dim mRec1 As New ADODB.Recordset
   Dim posInstr As Integer
   Dim qtyDecimales As Integer
   Dim mCodTipoVale As String
    
   mRet = True
      
   If Trim(Text3.Text) = "" Then
      mRet = False
      mMensajeError = "Debe completar el Número de Vale"
   End If

   If mRet Then
      If Not IsNumeric(Trim(Text3.Text)) Then
         mRet = False
         mMensajeError = "El Nro. Vale debe ser numérico !!"
      End If
   End If


   If mRet Then
      If Len(Trim(Text3.Text)) <> 9 Then
         mRet = False
         mMensajeError = "El Nro. Vale debe tener 9 caracteres numéricos !!"
      End If
   End If
   
   If mRet Then
      If ((Not Option1.Value) And (Not Option2.Value)) Then
         mRet = False
         mMensajeError = "Debe completar el Tipo de Vale"
      End If
   End If
         

   If mRet Then
      If Option1.Value Then
         mCodTipoVale = "C"
      Else
         mCodTipoVale = "M"
      End If

      Set mRec1 = mObj.oEjecutarSelect("SELECT NroVale,CodTipoVale FROM Consumos_H WHERE NroVale = " & Trim(Text3.Text) & " and CodTipoVale = '" & mCodTipoVale & "'; ")
      If Not mRec1.EOF Then
         mRet = False
         mMensajeError = "Ya se han realizado consumos para ese Número y Tipo de Vale !!!"
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
      For mj = 2 To FlexEgreso.Rows - 1
         If FlexEgreso.TextMatrix(mj, 6) = FlexProduct.TextMatrix(mRenglonProducto, 6) And FlexEgreso.TextMatrix(mj, 7) = FlexProduct.TextMatrix(mRenglonProducto, 7) Then
            mRet = False
            mMensajeError = "El Producto y la Ubicación elegidos ya han sido seleccionados"
            mj = 999
         End If
      Next
   End If
      
   'Valido si el saldo del stock es insuficiente para ese Producto/Ubicación
   If mRet Then
      
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
   
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   fValidaEgreso = mRet
End Function

Private Function fValidaConfirmarConsumo() As Boolean
   
   'Validaciones:
   'Que se haya completado el campo Numero Vale.
   'Que el Numero de Vale sea un valor entero.
   'Que se haya completado el combo "Retirado por:"
   'Que esten chequeados alguna de las dos radio button (Option1 u Option2)
   'Que en la grilla inferior "Egresos" exista al menos un registro.
   'Que no exista el Registro (Nro Vale, TipoVale ) en la tabla Consumo_H
 
   Dim mRet As Boolean
   Dim mMensajeError As String
   Dim mCodTipoVale As String
   Dim mRec1 As New ADODB.Recordset
   
   mRet = True
      
   If Trim(Text3.Text) = "" Then
      mRet = False
      mMensajeError = "Debe completar el Número de Vale"
   End If

   If mRet Then
      If Not IsNumeric(Trim(Text3.Text)) Then
         mRet = False
         mMensajeError = "El Nro. Vale debe ser numérico !!"
      End If
   End If
   
   If mRet Then
      If Len(Trim(Text3.Text)) <> 9 Then
         mRet = False
         mMensajeError = "El Nro. Vale debe tener 9 caracteres numéricos !!"
      End If
   End If
   
      
   If mRet Then
      If ((Not Option1.Value) And (Not Option2.Value)) Then
         mRet = False
         mMensajeError = "Debe completar el Tipo de Vale"
      End If
   End If
      
      
   If mRet Then
      If Trim(Right(Combo2.Text, 25)) = "" Then
         mRet = False
         mMensajeError = "Debe completar el campo: 'Retirado por'"
      End If
   End If
      
   If mRet Then
      If Trim(Right(Combo3.Text, 25)) = "" Then
         mRet = False
         mMensajeError = "Debe completar el campo: 'Autorizado por'"
      End If
   End If
      
    
   If mRet Then
      If FlexEgreso.Rows <= 2 Then
         mRet = False
         mMensajeError = "Al menos debe existir un registro en la Grilla Egresos"
      End If
   End If
   
   If mRet Then
      If Option1.Value Then
         mCodTipoVale = "C"
      Else
         mCodTipoVale = "M"
      End If

      Set mRec1 = mObj.oEjecutarSelect("SELECT NroVale,CodTipoVale FROM Consumos_H WHERE NroVale = " & Trim(Text3.Text) & " and CodTipoVale = '" & mCodTipoVale & "'; ")
      If Not mRec1.EOF Then
         mRet = False
         mMensajeError = "Ya se han realizado consumos para ese Número y Tipo de Vale !!!"
      End If
      mRec1.Close
   End If

   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   
   fValidaConfirmarConsumo = mRet
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

Private Sub FlexEgreso_Click()
   Dim mi As Integer
   
   If FlexEgreso.MouseRow > 0 Then
   
      If mRenglonEgreso <> 0 Then
         If FlexEgreso.Rows > mRenglonEgreso Then
            FlexEgreso.Row = mRenglonEgreso
            For mi = 1 To FlexEgreso.Cols - 1
               FlexEgreso.Col = mi
               FlexEgreso.CellBackColor = vbWhite
            Next
         End If
      End If
      
      mRenglonEgreso = FlexEgreso.MouseRow
   
      FlexEgreso.Row = mRenglonEgreso
      For mi = 1 To FlexEgreso.Cols - 1
         FlexEgreso.Col = mi
         FlexEgreso.CellBackColor = &H80000003
      Next
      
      If mRenglonEgreso > 1 Then
          mCodProducto = FlexEgreso.TextMatrix(mRenglonEgreso, 4)
      End If
   Else
      FlexEgreso.Row = mRenglonEgreso
      For mi = 1 To FlexProduct.Cols - 1
         FlexEgreso.Col = mi
         FlexEgreso.CellBackColor = vbWhite
      Next
      mRenglonEgreso = 0
   End If
End Sub
