VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form MantElect13 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingresos - Ajustes"
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
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   7440
      TabIndex        =   14
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton CommandProd 
      Height          =   375
      Index           =   1
      Left            =   8707
      Picture         =   "MantElect13.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4280
      Width           =   330
   End
   Begin VB.CommandButton CommandProd 
      Height          =   375
      Index           =   0
      Left            =   8047
      Picture         =   "MantElect13.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4280
      Width           =   330
   End
   Begin VB.Frame Frame1 
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
      Height          =   3565
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   16680
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   420
         Width           =   7335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   315
         Left            =   8760
         TabIndex        =   8
         Top             =   420
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid FlexProductDispo 
         Height          =   2460
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   16440
         _ExtentX        =   28998
         _ExtentY        =   4339
         _Version        =   327680
         Cols            =   8
      End
      Begin VB.Label Label3 
         Caption         =   "Contiene texto:"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Confirmar Ajuste"
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   6
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Index           =   1
      Left            =   9480
      TabIndex        =   5
      Top             =   8880
      Width           =   1815
   End
   Begin VB.Frame Frame10 
      Caption         =   "Ajustar Ingresos en Vale:"
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
      Height          =   4080
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   16680
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   278
         Left            =   13680
         TabIndex        =   3
         Top             =   7560
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid FlexProduct 
         Height          =   3480
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   16440
         _ExtentX        =   28998
         _ExtentY        =   6138
         _Version        =   327680
         Cols            =   9
      End
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Ingresado en:"
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
      Height          =   255
      Left            =   5760
      TabIndex        =   15
      Top             =   180
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Vale  - (Tipo Vale):"
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
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   165
      Width           =   2175
   End
End
Attribute VB_Name = "MantElect13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mObj As New clMantElect
Dim mObjInven As New clInven
Dim mRec As New ADODB.Recordset
Dim mRenglonProducto As Integer
Dim mRenglonProdOT As Integer
Dim mCodProducto As String

Dim filaAnt As Integer
Dim columnAnt As Integer

Dim mvMat_CodProd_Orig() As String
Dim mvMat_CodUbic_Orig() As String
Dim mvMat_Cantidad_Orig() As Double
Dim mvMat_CantidadBD_Orig() As Double
Private Sub Combo3_Click()
   Dim mi As Integer
   Dim mNroVale As String
   Dim mCodTipoVale As String
   Dim mCodUbicacion As String

   mNroVale = Left(Combo3.Text, 9)
   mCodTipoVale = Right(Combo3.Text, 1)
   
   Set mRec = mObj.oEjecutarSelect("SELECT I.NroVale, I.CodTipoVale,I.CodUbicacion, U.Descripcion AS Ubicacion " & _
   " FROM Mat_Ingresos_H I " & _
   " Inner Join " & _
   " Inventario.Ubicaciones U ON U.Codigo = I.CodUbicacion " & _
   " WHERE I.NroVale = '" & mNroVale & "' " & _
   " AND I.CodTipoVale ='" & mCodTipoVale & "';")
   
  Text1.Text = ""

   If Not mRec.EOF Then
      Text1 = mRec!Ubicacion & Space(100) & mRec!CodUbicacion
      mCodUbicacion = NVL(mRec!CodUbicacion, "")
   End If
   mRec.Close
   
   mRenglonProdOT = 0
   Text2.Text = ""
   Text2.Visible = False
   FlexProduct.ScrollBars = flexScrollBarVertical
   
   
   llenoGrillaAjustes mNroVale, mCodTipoVale, mCodUbicacion

   mRenglonProducto = 0
   'Elimino los registros (de la consulta anterior) de la grilla superior
   For mi = FlexProductDispo.Rows To 3 Step -1
      FlexProductDispo.RemoveItem mi
   Next

   preparaArrayMateriales mvMat_CodProd_Orig(), mvMat_CodUbic_Orig(), mvMat_Cantidad_Orig(), mvMat_CantidadBD_Orig()
   
End Sub


Private Sub llenoGrillaAjustes(pNroVale As String, pCodTipoVale As String, pCodUbicacion As String)
 Dim mi As Integer
 
 'Elimino los registros  de la grilla
  For mi = FlexProduct.Rows To 3 Step -1
      FlexProduct.RemoveItem mi
   Next
   
   'Set mRec = mObj.getConsumoAbastecimientoXidOTyUbicacion(pIdOT, pCodUbicacion)
   Set mRec = mObj.getIngresoXnroVale_CodTipoVale_Ubicacion(pNroVale, pCodTipoVale, pCodUbicacion)
   
   'Cargo la Grilla INFIOR con lo devuelto por el sp
   If Not mRec.EOF Then
      mi = 1
      Do While Not mRec.EOF
         mi = mi + 1

         FlexProduct.AddItem ""
         FlexProduct.TextMatrix(mi, 1) = mRec!CodigoSap
         FlexProduct.TextMatrix(mi, 2) = mRec!Producto
         FlexProduct.TextMatrix(mi, 3) = mRec!Cantidad
         FlexProduct.TextMatrix(mi, 4) = mRec!Stock
         FlexProduct.TextMatrix(mi, 5) = mRec!UnidadMedida
         FlexProduct.TextMatrix(mi, 6) = mRec!CodProducto
         FlexProduct.TextMatrix(mi, 7) = mRec!CodUbicacion
         FlexProduct.TextMatrix(mi, 8) = mRec!CantidadBD
         
         mRec.MoveNext
      Loop
   End If
   mRec.Close
End Sub

Private Sub Command1_Click()
   Dim mi As Integer
   Dim mj As Integer
   If Combo3.ListIndex <> -1 Then
      sMsgEspere Me, "Buscando productos...", True
      mRenglonProducto = 0
      'Elimino los registros (de la consulta anterior) de la grilla superior
      For mi = FlexProductDispo.Rows To 3 Step -1
         FlexProductDispo.RemoveItem mi
      Next
     
      'TODO: Dado que estamos consumiendo, seria ideal que el Store siguiente solo muestre los productos con stock > 0 en esa ubicacion.
      Set mRec = mObjInven.getStockXUbicacionConFiltroProducto(Right(Text1.Text, 4), Text4.Text)
      
      'Cargo la Grilla Superior con lo devuelto por el sp "getStockXBodegaConFiltroProducto"
      If Not mRec.EOF Then
         mi = 1
         Do While Not mRec.EOF
            mi = mi + 1
            
            FlexProductDispo.AddItem ""
            FlexProductDispo.TextMatrix(mi, 1) = mRec!CodigoSap
            FlexProductDispo.TextMatrix(mi, 2) = mRec!Producto
            FlexProductDispo.TextMatrix(mi, 3) = mRec!Ubicacion
            FlexProductDispo.TextMatrix(mi, 4) = mRec!Stock
            FlexProductDispo.TextMatrix(mi, 5) = mRec!UnidadMedida
            FlexProductDispo.TextMatrix(mi, 6) = mRec!CodProducto
            FlexProductDispo.TextMatrix(mi, 7) = mRec!CodUbicacion
            
            mRec.MoveNext
         Loop
      End If
      mRec.Close
      sMsgEspere Me, "", False
   Else
      MsgBox "Debe seleccionar algún Vale para poder buscar Productos.", vbExclamation, "Buscar Productos"
   End If
End Sub

Private Sub Command2_Click(Index As Integer)
   Dim vMat_CodProd_Ajuste() As String
   Dim vMat_CodUbic_Ajuste() As String
   Dim vMat_Cantidad_Ajuste() As Double
   Dim vMat_CantidadBD_Ajuste() As Double
   Dim mNroVale As String
   Dim mCodTipoVale As String

   If Index = 0 Then
      If fValidaConfirmarAjuste() Then
         preparaArrayMateriales vMat_CodProd_Ajuste(), vMat_CodUbic_Ajuste(), vMat_Cantidad_Ajuste(), vMat_CantidadBD_Ajuste()
         'fValidaConfirmarAjusteArray: Valido si es suficiente el stock para productos eliminados de la grilla inferior.
         If fValidaConfirmarAjusteArray(vMat_CodProd_Ajuste(), vMat_CodUbic_Ajuste(), vMat_Cantidad_Ajuste(), vMat_CantidadBD_Ajuste(), _
                                        mvMat_CodProd_Orig(), mvMat_CodUbic_Orig(), mvMat_Cantidad_Orig(), mvMat_CantidadBD_Orig()) Then
            mNroVale = Left(Combo3.Text, 9)
            mCodTipoVale = Right(Combo3.Text, 1)
      
            mObj.xAjustarIngreso mNroVale, mCodTipoVale, vMat_CodProd_Ajuste(), vMat_CodUbic_Ajuste(), vMat_Cantidad_Ajuste(), vMat_CantidadBD_Ajuste(), _
                                       mvMat_CodProd_Orig(), mvMat_CodUbic_Orig(), mvMat_Cantidad_Orig(), mvMat_CantidadBD_Orig(), Trim(Right(MDI.mUser, 15))
      
            Text2.Visible = False
           llenoGrillaAjustes mNroVale, mCodTipoVale, Right(Trim(Text1.Text), 4)
           preparaArrayMateriales mvMat_CodProd_Orig(), mvMat_CodUbic_Orig(), mvMat_Cantidad_Orig(), mvMat_CantidadBD_Orig()
            
            MsgBox "Se ha ajustado correctamente los materiales del Ingreso ", vbInformation
         End If
      End If
   Else
      Unload Me
   End If
End Sub

Private Sub CommandProd_Click(Index As Integer)
   
   Dim mRec1 As New ADODB.Recordset
   Dim iStock As Double
   
   If Index = 0 Then
      If fValidar() Then
         FlexProduct.AddItem vbTab & FlexProductDispo.TextMatrix(mRenglonProducto, 1) & vbTab & FlexProductDispo.TextMatrix(mRenglonProducto, 2) & vbTab & getCantidadEnVectorProductOrig(FlexProductDispo.TextMatrix(mRenglonProducto, 6)) & vbTab & FlexProductDispo.TextMatrix(mRenglonProducto, 4) & vbTab & FlexProductDispo.TextMatrix(mRenglonProducto, 5) & vbTab & FlexProductDispo.TextMatrix(mRenglonProducto, 6) & vbTab & FlexProductDispo.TextMatrix(mRenglonProducto, 7) & vbTab & getCantidadEnVectorProductOrig(FlexProductDispo.TextMatrix(mRenglonProducto, 6))
      End If
   Else
      If FlexProduct.Rows > 2 And mRenglonProdOT > 1 Then
         
         
         
         Set mRec1 = mObj.oEjecutarSelect("SELECT Fecha,CodProducto,CodUbicacion, IFNULL(Stock,0) As Stock " & _
                                          " FROM Inventario.Movimientos2 M " & _
                                          " WHERE CodProducto  = '" & FlexProduct.TextMatrix(mRenglonProdOT, 6) & "' and CodUbicacion = '" & FlexProduct.TextMatrix(mRenglonProdOT, 7) & "'" & _
                                          " AND Fecha = (SELECT Max(Fecha) FROM Inventario.Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion)")
         If Not mRec1.EOF Then
            iStock = mRec1!Stock
         Else
            iStock = 0
         End If
         mRec1.Close
         
            
         'Valido si cuando elimino el registro en la grilla inferior, tenga suficiente stock porque ello va a significar un Egreso.
         'If CDbl(Replace(Trim(FlexProduct.TextMatrix(mRenglonProdOT, 8)), ".", ",")) > CDbl(Replace(Trim(FlexProduct.TextMatrix(mRenglonProdOT, 4)), ".", ",")) Then
         If CDbl(Replace(Trim(FlexProduct.TextMatrix(mRenglonProdOT, 8)), ".", ",")) > iStock Then
             MsgBox "Stock Insuficiente !!!" & Chr(13) & "PRODUCTO: " & Trim(FlexProduct.TextMatrix(mRenglonProdOT, 2)) & Chr(13) & "No es posible realizar una salida por " & Trim(FlexProduct.TextMatrix(mRenglonProdOT, 8)) & " " & Trim(FlexProduct.TextMatrix(mRenglonProdOT, 5)), vbCritical, "Atención"
         Else
            FlexProduct.RemoveItem (mRenglonProdOT)
            filaAnt = 0
            columnAnt = 0
            Text2.Visible = False
            mRenglonProdOT = 0
         End If
      
      End If
      
   End If
End Sub

Private Function getCantidadEnVectorProductOrig(pCodProducto As String) As Double
Dim ret As Double
Dim mi As Integer

ret = 0
For mi = LBound(mvMat_CodProd_Orig) To UBound(mvMat_CodProd_Orig)
   If pCodProducto = mvMat_CodProd_Orig(mi) Then
      ret = mvMat_Cantidad_Orig(mi)
      mi = 9999
   End If
Next
getCantidadEnVectorProductOrig = ret
End Function
Private Function fValidar() As Boolean
   Dim mRet As Boolean
   Dim mMensajeError As String
   Dim mj As Integer
   
   mRet = True
   
   If mRenglonProducto = 0 Then
      mRet = False
      mMensajeError = "Debe seleccionar un Producto de la Grilla superior"
   End If
      
   If mRet Then
      If mRenglonProducto <> 0 And FlexProductDispo.TextMatrix(mRenglonProducto, 6) = "" Then
         mRet = False
         mMensajeError = "Debe seleccionar un Producto de la Grilla superior"
      End If
   End If
      
   'Valido si el Producto y la Ubicacion seleccionados coinciden con un Egreso ya registrado
   If mRet Then
      For mj = 2 To FlexProduct.Rows - 1
         If FlexProduct.TextMatrix(mj, 6) = FlexProductDispo.TextMatrix(mRenglonProducto, 6) And FlexProduct.TextMatrix(mj, 7) = FlexProductDispo.TextMatrix(mRenglonProducto, 7) Then
            mRet = False
            mMensajeError = "El Producto y la Ubicación elegidos ya se encuentran en la grilla inferior"
            mj = 999
         End If
      Next
   End If
      
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   fValidar = mRet
End Function

Private Function fValidaConfirmarAjuste() As Boolean
   Dim mRet As Boolean
   Dim mMensajeError As String
   Dim mRec1 As New ADODB.Recordset
   Dim mi As Integer
   Dim posInstr As Integer
   Dim qtyDecimales As Integer
   Dim iStock As Double

   mRet = True
   
   'Valido Cantidad valida, cantidad decimales <2 t  saldo del stock insuficiente para ese Producto/Ubicación
   If mRet Then
      If mRet Then
         For mi = 2 To FlexProduct.Rows - 1
            Set mRec1 = mObjInven.oEjecutarSelect("SELECT Fecha,CodProducto,CodUbicacion, IFNULL(Stock,0) As Stock " & _
                                                      " FROM Movimientos2 M " & _
                                                      " WHERE CodProducto  = '" & FlexProduct.TextMatrix(mi, 6) & "' and CodUbicacion = '" & FlexProduct.TextMatrix(mi, 7) & "'" & _
                                                      " AND Fecha = (SELECT Max(Fecha) FROM Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion)")
            If Not mRec1.EOF Then
               iStock = mRec1!Stock
            Else
               iStock = 0
            End If
            mRec1.Close
            
            posInstr = InStr(1, Replace(FlexProduct.TextMatrix(mi, 3), ".", ","), ",")
      
            qtyDecimales = 0
            If posInstr <> 0 Then
               qtyDecimales = Len(Right(Trim(FlexProduct.TextMatrix(mi, 3)), Len(Trim(FlexProduct.TextMatrix(mi, 3))) - posInstr))
            End If
            
            
            'Valido valor numerico
            If Not IsNumeric(Replace(FlexProduct.TextMatrix(mi, 3), ".", ",")) Then
               mRet = False
               mMensajeError = "Se ha cargado un valor incorrecto para el producto: '" & FlexProduct.TextMatrix(mi, 2) & "'"
               mi = 9999
            End If
            
            'Valido cantidad decimales
            If mRet Then
               If qtyDecimales > 2 Then
                  mRet = False
                   mMensajeError = "La Cantidad ingresada para  ' " & FlexProduct.TextMatrix(mi, 2) & " ' no puede tener mas de dos decimales"
                  mi = 9999
               End If
            End If
            
            If mRet Then
               If CDbl(Replace(Trim(FlexProduct.TextMatrix(mi, 3)), ".", ",")) < 0 Then
                  mRet = False
                  mMensajeError = "Se ha cargado un valor menor a cero para el producto: '" & FlexProduct.TextMatrix(mi, 2) & "'"
                  mi = 9999
               End If
            End If
            'Valido saldo insuficiente (En caso que implique una salida de material)
            If mRet Then
               'Si es salida de material
               If CDbl(Replace(Trim(FlexProduct.TextMatrix(mi, 3)), ".", ",")) < CDbl(Replace(Trim(FlexProduct.TextMatrix(mi, 8)), ".", ",")) Then
                  If (CDbl(Replace(Trim(FlexProduct.TextMatrix(mi, 8)), ".", ",")) - CDbl(Replace(Trim(FlexProduct.TextMatrix(mi, 3)), ".", ","))) > iStock Then
                     mRet = False
                     mMensajeError = "El stock es insuficiente para ' " & FlexProduct.TextMatrix(mi, 2) & " '"
                     mi = 9999
                  End If
               End If
               
            End If
         Next
      End If
   End If
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If

   fValidaConfirmarAjuste = mRet
End Function

Public Function fValidaConfirmarAjusteArray(ByRef pvMat_CodProd_Ajuste() As String, ByRef pvMat_CodUbic_Ajuste() As String, _
                                             ByRef pvMat_Cantidad_Ajuste() As Double, ByRef pvMat_CantidadBD_Ajuste() As Double, _
                                             ByRef pvMat_CodProd_Orig() As String, ByRef pvMat_CodUbic_Orig() As String, _
                                             ByRef pvMat_Cantidad_Orig() As Double, ByRef pvMat_CantidadBD_Orig() As Double) As Boolean
   Dim mi As Integer
   Dim mRet As Boolean
   Dim mMensajeError As String
   Dim mRec1 As New ADODB.Recordset
   Dim iStock As Double
   Dim ProdDescr As String
   
   mRet = True
   
   If pvMat_CodProd_Ajuste(0) <> "000000" Then

      For mi = LBound(pvMat_CodProd_Orig) To UBound(pvMat_CodProd_Orig)
         If Not mObj.estaEnVectorAjuste(pvMat_CodProd_Orig(mi), pvMat_CodProd_Ajuste()) Then
            
            Set mRec1 = mObj.oEjecutarSelect("SELECT Fecha,CodProducto,CodUbicacion, IFNULL(Stock,0) As Stock " & _
                                             " FROM Inventario.Movimientos2 M " & _
                                             " WHERE CodProducto  = '" & pvMat_CodProd_Orig(mi) & "' and CodUbicacion = '" & pvMat_CodUbic_Orig(mi) & "'" & _
                                             " AND Fecha = (SELECT Max(Fecha) FROM Inventario.Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion)")
            If Not mRec1.EOF Then
               iStock = mRec1!Stock
            Else
               iStock = 0
            End If
            mRec1.Close
            
            If pvMat_CantidadBD_Orig(mi) > iStock Then
               ProdDescr = mObj.sCampoDescrip("Inventario.Producto", "Codigo='" & pvMat_CodProd_Orig(mi) & "'", 1)
               mRet = False
               mMensajeError = "El stock es insuficiente para ' " & ProdDescr & " '"
               mi = 9999
            End If

         End If
      Next
   End If
   
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   
   fValidaConfirmarAjusteArray = mRet
End Function


Private Sub FlexProduct_Click()
   Dim mi As Integer
   
   If FlexProduct.MouseRow > 0 Then
      'If Not mEsOTcerrada Then
         'En este caso 3 es la columna que seria editable
         If FlexProduct.Col = 3 And FlexProduct.Row <> 1 Then
            Text2.Text = FlexProduct.Text
            Text2.Width = FlexProduct.ColWidth(FlexProduct.Col)
            Text2.Left = FlexProduct.ColPos(FlexProduct.Col) + FlexProduct.Left + 30 'el valor treina termina de acomodar el textbox en la celda
            Text2.Top = FlexProduct.Top + FlexProduct.RowPos(FlexProduct.Row)
            Text2.Visible = True
            Text2.SetFocus
            FlexProduct.ScrollBars = flexScrollBarNone
         Else
            Text2.Visible = False
            FlexProduct.ScrollBars = flexScrollBarVertical
         End If
      
         filaAnt = FlexProduct.Row
         columnAnt = FlexProduct.Col
      'End If
      If mRenglonProdOT <> 0 Then
         FlexProduct.Row = mRenglonProdOT
         For mi = 1 To FlexProduct.Cols - 1
            FlexProduct.Col = mi
            FlexProduct.CellBackColor = vbWhite
         Next
      End If
      mRenglonProdOT = FlexProduct.MouseRow
      FlexProduct.Row = mRenglonProdOT
      For mi = 1 To FlexProduct.Cols - 1
         FlexProduct.Col = mi
         FlexProduct.CellBackColor = &H80000003
      Next
      If mRenglonProdOT > 1 Then
          mCodProducto = FlexProduct.TextMatrix(mRenglonProdOT, 4)
      End If
   Else
      FlexProduct.Row = mRenglonProdOT
      If FlexProduct.Row > 0 Then
         For mi = 1 To FlexProduct.Cols - 1
            FlexProduct.Col = mi
            FlexProduct.CellBackColor = vbWhite
         Next
      End If
      mRenglonProdOT = 0
   End If

End Sub

'Private Sub FlexProductDispoDispo_Click()
'   Dim mi As Integer
'   If FlexProductDispo.MouseRow > 0 Then
'
'      If mRenglonProducto <> 0 Then
'         FlexProductDispo.Row = mRenglonProducto
'         For mi = 1 To FlexProductDispo.Cols - 1
'            FlexProductDispo.Col = mi
'            FlexProductDispo.CellBackColor = vbWhite
'         Next
'      End If
'      mRenglonProducto = FlexProductDispo.MouseRow
'      FlexProductDispo.Row = mRenglonProducto
'      For mi = 1 To FlexProductDispo.Cols - 1
'         FlexProductDispo.Col = mi
'         FlexProductDispo.CellBackColor = &H80000003
'      Next
'      If mRenglonProducto > 1 Then
'          mCodProducto = FlexProductDispo.TextMatrix(mRenglonProducto, 6)
'      End If
'   Else
'      FlexProductDispo.Row = mRenglonProducto
'      For mi = 1 To FlexProductDispo.Cols - 1
'         FlexProductDispo.Col = mi
'         FlexProductDispo.CellBackColor = vbWhite
'      Next
'      mRenglo
'End Sub

Private Sub FlexProductDispo_Click()
   Dim mi As Integer
   If FlexProductDispo.MouseRow > 0 Then
   
      If mRenglonProducto <> 0 Then
         FlexProductDispo.Row = mRenglonProducto
         For mi = 1 To FlexProductDispo.Cols - 1
            FlexProductDispo.Col = mi
            FlexProductDispo.CellBackColor = vbWhite
         Next
      End If
      mRenglonProducto = FlexProductDispo.MouseRow
      FlexProductDispo.Row = mRenglonProducto
      For mi = 1 To FlexProductDispo.Cols - 1
         FlexProductDispo.Col = mi
         FlexProductDispo.CellBackColor = &H80000003
      Next
      If mRenglonProducto > 1 Then
          mCodProducto = FlexProductDispo.TextMatrix(mRenglonProducto, 6)
      End If
   Else
      FlexProductDispo.Row = mRenglonProducto
      For mi = 1 To FlexProductDispo.Cols - 1
         FlexProductDispo.Col = mi
         FlexProductDispo.CellBackColor = vbWhite
      Next
      mRenglonProducto = 0
   End If
End Sub

Private Sub Form_Load()
   Me.Width = 17085
   Me.Height = 9920
   sAlinearForm Me
   
   Text1.Enabled = False
   
   
   Set mRec = mObj.oEjecutarSelect("SELECT CONCAT(NroVale,' - (', CASE WHEN CodTipoVale = 'C' THEN 'Vale a Cargo' ELSE 'Vale Retiro Múltiple' END, ')') AS Vale_TipoVale, Fecha, CodTipoVale " & _
                           " From Mat_Ingresos_H " & _
                           " ORDER BY NroVale; ")
                           
   Do While Not mRec.EOF
      Combo3.AddItem mRec!Vale_TipoVale & Space(100) & mRec!CodTipoVale
      mRec.MoveNext
   Loop
   mRec.Close
   initMateriales
End Sub

Private Sub initMateriales()
   filaAnt = 0
   columnAnt = 0
   Text2.Visible = False
   
   With FlexProductDispo
      .ColWidth(0) = 200
      .ColWidth(1) = 950
      .ColWidth(2) = 9700
      .ColWidth(3) = 2100
      .ColWidth(4) = 1200
      .ColWidth(5) = 1900
      
      .ColWidth(6) = 0
      .ColWidth(7) = 0
      
      .TextMatrix(0, 1) = "Cód.Sap"
      .TextMatrix(0, 2) = "Producto"
      .TextMatrix(0, 3) = "Ubicación"
      .TextMatrix(0, 4) = "Stock"
      .TextMatrix(0, 5) = "Unid.Medida"
      
      .TextMatrix(0, 6) = "Cód. Producto"
      .TextMatrix(0, 7) = "Cód. Ubicacion"
      
      .RowHeight(1) = 0
   End With
   
   With FlexProduct
      .ColWidth(0) = 200
      .ColWidth(1) = 950
      .ColWidth(2) = 9700
      .ColWidth(3) = 1650
      .ColWidth(4) = 1650
      .ColWidth(5) = 1900
      .ColWidth(6) = 0
      .ColWidth(7) = 0
      .ColWidth(8) = 0
      
      .TextMatrix(0, 1) = "Cód.Sap"
      .TextMatrix(0, 2) = "Producto"
      .TextMatrix(0, 3) = "Cantidad"
      .TextMatrix(0, 4) = "Stock"
      .TextMatrix(0, 5) = "Unid.Medida"
      .TextMatrix(0, 6) = "Cód. Producto"
      .TextMatrix(0, 7) = "Cód. Ubicacion"
      .TextMatrix(0, 8) = "CantidadBD"
      
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
      FlexProduct.TextMatrix(filaAnt, columnAnt) = Text2.Text
      Text2.Visible = False
      FlexProduct.ScrollBars = flexScrollBarVertical
   End If
End Sub

Private Sub Text2_LostFocus()
   If FlexProduct.Col <> columnAnt Or FlexProduct.Row <> filaAnt Then
      'En este caso 3 es la columna que seria editable
      If columnAnt = 3 Then
         FlexProduct.TextMatrix(filaAnt, columnAnt) = Text2.Text
      End If
   End If
End Sub

Private Sub preparaArrayMateriales(ByRef pvMat_CodProd() As String, ByRef pvMat_CodUbic() As String, ByRef pvMat_Cantidad() As Double, ByRef pvMat_CantidadBD() As Double)
   Dim mj As Integer
   Dim cantMateriales As Integer

   cantMateriales = FlexProduct.Rows - 2
   If cantMateriales > 0 Then
      
      ReDim pvMat_CodProd(0 To cantMateriales - 1) As String
      ReDim pvMat_CodUbic(0 To cantMateriales - 1) As String
      ReDim pvMat_Cantidad(0 To cantMateriales - 1) As Double
      ReDim pvMat_CantidadBD(0 To cantMateriales - 1) As Double
      
      For mj = 2 To FlexProduct.Rows - 1
        pvMat_CodProd(mj - 2) = FlexProduct.TextMatrix(mj, 6)
        pvMat_CodUbic(mj - 2) = FlexProduct.TextMatrix(mj, 7)
        pvMat_Cantidad(mj - 2) = CDbl(Replace(FlexProduct.TextMatrix(mj, 3), ".", ","))
        pvMat_CantidadBD(mj - 2) = CDbl(Replace(FlexProduct.TextMatrix(mj, 8), ".", ","))
      Next
   Else
      'Esta linea sirve como flag para avisar en el procedimiento xinsOT que no tiene registros y por ende no me de error de intervalos al usar Ubound y Lubound
      ReDim pvMat_CodProd(0)
      pvMat_CodProd(0) = "000000"
   End If
End Sub


