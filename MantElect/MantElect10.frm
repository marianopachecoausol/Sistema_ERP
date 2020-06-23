VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form MantElect10 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Materiales"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16965
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   16965
   ShowInTaskbar   =   0   'False
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
      Left            =   9210
      TabIndex        =   18
      Top             =   8600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Confirmar"
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
      Left            =   6540
      TabIndex        =   17
      Top             =   8600
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   960
      TabIndex        =   15
      Top             =   4120
      Width           =   1575
   End
   Begin VB.CommandButton CommandProd 
      Height          =   375
      Index           =   0
      Left            =   2760
      Picture         =   "MantElect10.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4080
      Width           =   330
   End
   Begin VB.CommandButton CommandProd 
      Height          =   375
      Index           =   1
      Left            =   3240
      Picture         =   "MantElect10.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4080
      Width           =   330
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Vale de retiro múltiple"
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
      Height          =   375
      Left            =   6960
      TabIndex        =   12
      Top             =   60
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Vale a cargo/recambio       /"
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
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   60
      Width           =   3255
   End
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
      Left            =   2040
      MaxLength       =   9
      TabIndex        =   10
      Top             =   60
      Width           =   1215
   End
   Begin VB.Frame Frame11 
      Caption         =   "Ingresos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   3780
      Left            =   120
      TabIndex        =   7
      Top             =   4600
      Width           =   16680
      Begin MSFlexGridLib.MSFlexGrid FlexIngreso 
         Height          =   3255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   16440
         _ExtentX        =   28998
         _ExtentY        =   5741
         _Version        =   327680
         Cols            =   9
      End
   End
   Begin VB.Frame Frame10 
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
      TabIndex        =   0
      Top             =   600
      Width           =   16680
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   315
         Left            =   15000
         TabIndex        =   3
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   7560
         TabIndex        =   2
         Top             =   420
         Width           =   7335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   420
         Width           =   3735
      End
      Begin MSFlexGridLib.MSFlexGrid FlexProduct 
         Height          =   2295
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   16440
         _ExtentX        =   28998
         _ExtentY        =   4048
         _Version        =   327680
         Cols            =   8
      End
      Begin VB.Label Label3 
         Caption         =   "Contiene texto:"
         Height          =   375
         Left            =   6360
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Ingresar en:"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Vale número:"
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
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "MantElect10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mObj As New clMantElect
Dim mObjInven As New clInven
Dim mRec As New ADODB.Recordset

Dim mRenglonProducto As Integer
Dim mRenglonIngreso As Integer
Dim mCodProducto As String

Dim mCodTipoVale As String
Dim mOTtieneValeAsociado As Boolean

Private Sub Combo3_Click()
'   Dim IdOT As Integer
'   Dim mi As Integer
'   Dim mResultado As Boolean
'
   'IdOT = CInt(Left(Combo3.Text, 10))

   'Set mRec = mObj.oEjecutarSelect("SELECT IdOT, NroVale,CodTipoVale FROM OT_Abastecimiento_H O Where IdOT = " & IdOT & "; ")
   
'   If Not mRec.EOF Then
'      mOTtieneValeAsociado = True
'      Text3.Text = mRec!NroVale
'      Text3.Enabled = False
'
'      If mRec!CodTipoVale = "C" Then
'         Option1.Value = True
'         Option2.Value = False
'      Else
'         Option1.Value = False
'         Option2.Value = True
'      End If
'      Option1.Enabled = False
'      Option2.Enabled = False
'
'      Command3(0).Enabled = False
'      CommandProd(0).Enabled = False
'      CommandProd(1).Enabled = False
'   Else
'      mOTtieneValeAsociado = False
'      Text3.Enabled = True
'      Text3.Text = ""
'      Option1.Value = False
'      Option2.Value = False
'      Option1.Enabled = True
'      Option2.Enabled = True
'      Command3(0).Enabled = True
'      CommandProd(0).Enabled = True
'      CommandProd(1).Enabled = True
'   End If
'   mRec.Close
   
'   For mi = FlexIngreso.Rows To 3 Step -1
'      FlexIngreso.RemoveItem mi
'   Next
'
'   If mOTtieneValeAsociado Then
'
'
'
'
'   Set mRec = mObj.oEjecutarSelect(" SELECT P.CodigoSap, P.Descripcion AS Producto, U.Descripcion AS Ubicacion, M.Cantidad, " & _
'                                   " UM.Descripcion as UnidadMedida,M.CodProducto,M.CodUbicacion FROM " & _
'                                   " OT_Abastecimiento_Det A " & _
'                                   "   INNER JOIN " & _
'                                   " Inventario.Movimientos2 M ON A.IdMov = M.IdMov " & _
'                                   "  INNER JOIN " & _
'                                   " Inventario.Producto P ON P.Codigo = M.CodProducto " & _
'                                   "  INNER JOIN " & _
'                                   " Inventario.Ubicaciones U ON U.Codigo = M.CodUbicacion " & _
'                                   "  INNER JOIN " & _
'                                   "   Inventario.UnidadMedida UM ON UM.Codigo = P.CodUnidadMedida " & _
'                                   " WHERE esMovimientoFinal = 1 " & _
'                                   " AND IdOT = " & IdOT & ";")
'
'
'
'
'      If Not mRec.EOF Then
'         mi = 1
'         Do While Not mRec.EOF
'            mi = mi + 1
'
'            With FlexIngreso
'               .AddItem ""
'               .TextMatrix(mi, 1) = mRec!CodigoSap
'               .TextMatrix(mi, 2) = mRec!Producto
'               .TextMatrix(mi, 3) = mRec!Ubicacion
'               .TextMatrix(mi, 4) = mRec!Cantidad
'               .TextMatrix(mi, 5) = mRec!UnidadMedida
'               .TextMatrix(mi, 6) = mRec!CodProducto
'               .TextMatrix(mi, 7) = mRec!CodUbicacion
'            End With
'
'            mRec.MoveNext
'         Loop
'      End If
'
'      mRec.Close
'   End If

End Sub



Private Sub Command1_Click()
   Dim mi As Integer
   Dim mj As Integer
   sMsgEspere Me, "Buscando productos...", True
   mRenglonProducto = 0
   
   'Elimino los registros de la grilla superior
   For mi = FlexProduct.Rows To 3 Step -1
      FlexProduct.RemoveItem mi
   Next
   'TODO: Dado que estamos consumiendo, seria ideal que el Store siguiente solo muestre los productos con stock > 0 en esa ubicacion.
   Set mRec = mObjInven.getStockXUbicacionConFiltroProducto(Right(Combo2.Text, 4), Text1.Text)
   
   'Cargo la Grilla Superior con lo devuelto por el sp "getStockXBodegaConFiltroProducto"
   If Not mRec.EOF Then
      mi = 1
      Do While Not mRec.EOF
         mi = mi + 1
         
         FlexProduct.AddItem ""
         FlexProduct.TextMatrix(mi, 1) = mRec!CodigoSap
         FlexProduct.TextMatrix(mi, 2) = mRec!Producto
         FlexProduct.TextMatrix(mi, 3) = mRec!Ubicacion
         FlexProduct.TextMatrix(mi, 4) = mRec!Stock
         FlexProduct.TextMatrix(mi, 5) = mRec!UnidadMedida
         FlexProduct.TextMatrix(mi, 6) = mRec!CodProducto
         FlexProduct.TextMatrix(mi, 7) = mRec!CodUbicacion
         
         mRec.MoveNext
      Loop
   End If
   mRec.Close
   
   'Si presiono el boton Buscar y algun "producto/ubicacion" de la grilla de arriba, esta en la grilla inferior
   'entonces que me actualice en la grilla superior el stock de ese o esos "producto/ubicacion" descontando el consumo de la grilla inferior
'   For mi = 2 To FlexProduct.Rows - 1
'      For mJ = 2 To FlexIngreso.Rows - 1
'         If FlexProduct.TextMatrix(mi, 6) = FlexIngreso.TextMatrix(mJ, 6) And FlexProduct.TextMatrix(mi, 7) = FlexIngreso.TextMatrix(mJ, 7) Then
'            FlexProduct.TextMatrix(mi, 4) = CDbl(Replace(Trim(FlexProduct.TextMatrix(mi, 4)), ".", ",")) - CDbl(Replace(Trim(FlexIngreso.TextMatrix(mJ, 4)), ".", ","))
'            mJ = 999
'         End If
'      Next
'   Next
   sMsgEspere Me, "", False
End Sub

Private Sub Command3_Click(Index As Integer)
If Index = 0 Then
   Dim vIngresosCodProducto() As String
   Dim vIngresosCodUbicacion() As String
   Dim vIngresosCantidad() As Double
   Dim vIngresosCantidadBD() As Double
   Dim cantIngresos As Integer
   Dim mj As Integer
   Dim mCodTipoVale As String
   Dim CodUbicacion As String
   Dim mResultado As Boolean


    CodUbicacion = Right(Trim(Combo2.Text), 4)

   If fValidaConfirmarIngreso() Then
      cantIngresos = FlexIngreso.Rows - 2
   
      ReDim vIngresosCodProducto(0 To cantIngresos - 1) As String
      ReDim vIngresosCodUbicacion(0 To cantIngresos - 1) As String
      ReDim vIngresosCantidad(0 To cantIngresos - 1) As Double
      ReDim vIngresosCantidadBD(0 To cantIngresos - 1) As Double
      
      For mj = 2 To FlexIngreso.Rows - 1
         vIngresosCodProducto(mj - 2) = FlexIngreso.TextMatrix(mj, 6)
         vIngresosCodUbicacion(mj - 2) = FlexIngreso.TextMatrix(mj, 7)
         vIngresosCantidad(mj - 2) = CDbl(Replace(FlexIngreso.TextMatrix(mj, 4), ".", ","))
         vIngresosCantidadBD(mj - 2) = CDbl(Replace(FlexIngreso.TextMatrix(mj, 8), ".", ","))
      Next
      
      If Option1.Value Then
         mCodTipoVale = "C"
      Else
         mCodTipoVale = "M"
      End If
      
      mResultado = True
      mObj.xInsMat_Ingresos vIngresosCodProducto(), vIngresosCodUbicacion(), vIngresosCantidad(), vIngresosCantidadBD(), Trim(Text3.Text), mCodTipoVale, CodUbicacion, Trim(Right(MDI.mUser, 15)), mResultado

      If mResultado Then
         
'         mOTtieneValeAsociado = True
'         Command3(0).Enabled = False
'         CommandProd(0).Enabled = False
'         CommandProd(1).Enabled = False
'         Text3.Enabled = False
'         Option1.Enabled = False
'         Option2.Enabled = False
         
         MsgBox "El Ingreso se ha realizado exitosamente", vbInformation, "Ingresos"
         limpioFormulario
         
         'VerificaStockMin vEgresosCodProducto(), mCodBodega
      End If
   End If
Else
   Unload Me
End If

End Sub

Private Sub limpioFormulario()
   Dim mi As Integer

   Text1.Text = ""
   Text2.Text = ""
   Text3.Text = ""
   Option1.Value = False
   Option2.Value = False
   
  
   'Elimino los registros de la grilla superior
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




Private Function fValidaConfirmarIngreso() As Boolean
  'Validado: 'Que se haya completado el campo Numero Vale.
  'Validado: 'Que el Numero de Vale sea un valor entero.
  'Validado:  'Que esten chequeados alguna de las dos radio button (Option1 u Option2)
  'Validado: 'Que en la grilla inferior "Ingresos" exista al menos un registro.
  'Validado: 'Que no exista el Registro (Nro Vale, TipoVale ) en la tabla Mat_Ingresos_H
 
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
      If FlexIngreso.Rows <= 2 Then
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
  
      Set mRec1 = mObj.oEjecutarSelect("SELECT NroVale,CodTipoVale FROM Mat_Ingresos_H WHERE NroVale = " & Trim(Text3.Text) & " and CodTipoVale = '" & mCodTipoVale & "'; ")
      If Not mRec1.EOF Then
         mRet = False
         mMensajeError = "Ya se han realizado ingresos para ese Número y Tipo de Vale !!!"
      End If
      mRec1.Close
   End If

   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   
   fValidaConfirmarIngreso = mRet
End Function


Private Sub CommandProd_Click(Index As Integer)
Dim iStock As Double
   Dim mi As Integer
   Dim mRec1 As New ADODB.Recordset
   
   If Index = 0 Then
      'TODO: Completar la funcion fvalidaEgreso
      If fValidaIngreso() Then
         FlexIngreso.AddItem vbTab & FlexProduct.TextMatrix(mRenglonProducto, 1) & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 2) & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 3) & vbTab & Text2.Text & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 5) & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 6) & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 7) & vbTab & 0
         'FlexProduct.TextMatrix(mRenglonProducto, 4) = Format(CDbl(Replace(Trim(FlexProduct.TextMatrix(mRenglonProducto, 4)), ".", ",")) - CDbl(Replace(Trim(Text2.Text), ".", ",")), "00.00")
         Text2.Text = ""
         Text2.SetFocus
      End If
   Else
'      For mi = 2 To FlexProduct.Rows - 1
'
'         If FlexProduct.TextMatrix(mi, 6) = FlexIngreso.TextMatrix(mRenglonIngreso, 6) And FlexProduct.TextMatrix(mi, 7) = FlexIngreso.TextMatrix(mRenglonIngreso, 7) Then
'            Set mRec1 = mObj.oEjecutarSelect("SELECT Fecha,CodProducto,CodUbicacion, IFNULL(Stock,0) As Stock " & _
'                                       " FROM Inventario.Movimientos2 M " & _
'                                       " WHERE CodProducto  = '" & FlexIngreso.TextMatrix(mRenglonIngreso, 6) & "' and CodUbicacion = '" & FlexIngreso.TextMatrix(mRenglonIngreso, 7) & "'" & _
'                                       " AND Fecha = (SELECT Max(Fecha) FROM Inventario.Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion)")
'
'            If Not mRec1.EOF Then
'               iStock = mRec1!Stock
'            Else
'               iStock = 0
'            End If
'            mRec1.Close
'
'            FlexProduct.TextMatrix(mi, 4) = iStock
'
'            mi = 9999
'         End If
'      Next
'
      If FlexIngreso.Rows > 2 And mRenglonIngreso > 1 Then
         FlexIngreso.RemoveItem (mRenglonIngreso)
      End If
      
      mRenglonIngreso = 0
   End If

End Sub

Private Sub FlexIngreso_Click()
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
          mCodProducto = FlexProduct.TextMatrix(mRenglonProducto, 6)
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
   Me.Width = 17090
   Me.Height = 9750
   sAlinearForm Me
   
'   Command3(0).Enabled = False
'   CommandProd(0).Enabled = False
'   CommandProd(1).Enabled = False
   
   
'   Set mRec = mObj.oEjecutarSelect("SELECT CONVERT( CONCAT(LPAD(IdOT,10,'0'),' - ',Date_Format(Fecha,'%d/%m/%Y')),char(23)) AS OT_Fecha, O.Fecha " & _
'                           " FROM MantElect.OT_H O " & _
'                           " ORDER BY IdOT DESC; ")
'
'   Do While Not mRec.EOF
'      Combo3.AddItem mRec!OT_Fecha
'      mRec.MoveNext
'   Loop
'   mRec.Close
   
   initMateriales
End Sub

Private Sub initMateriales()
   With FlexProduct
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

   With FlexIngreso
      .ColWidth(0) = 200
      .ColWidth(1) = 950
      .ColWidth(2) = 9700
      .ColWidth(3) = 2100
      .ColWidth(4) = 1200
      .ColWidth(5) = 1900
    
      .ColWidth(6) = 0
      .ColWidth(7) = 0
      .ColWidth(8) = 0

   
      .TextMatrix(0, 1) = "Cód.Sap"
      .TextMatrix(0, 2) = "Producto"
      .TextMatrix(0, 3) = "Ubicación"
      .TextMatrix(0, 4) = "Cantidad"
      .TextMatrix(0, 5) = "Unid.Medida"
      .TextMatrix(0, 6) = "Cód. Producto"
      .TextMatrix(0, 7) = "Cód. Ubicacion"
      .TextMatrix(0, 8) = "CantidadBD"


      .RowHeight(1) = 0
   End With

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

Private Sub Form_Unload(Cancel As Integer)
   Set mObj = Nothing
   Set mRec = Nothing
   ShowMenu 47, True, False
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
      If KeyAscii <> 46 Then
         KeyAscii = fNumeroKeyPress(KeyAscii)
      End If
End Sub

Private Function fValidaIngreso() As Boolean
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
      Set mRec1 = mObj.oEjecutarSelect("SELECT NroVale,CodTipoVale FROM Mat_Ingresos_H WHERE NroVale = " & Trim(Text3.Text) & " and CodTipoVale = '" & mCodTipoVale & "'; ")
      If Not mRec1.EOF Then
         mRet = False
         mMensajeError = "Ya se han realizado ingresos para ese Número y Tipo de Vale !!!"
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
      If mRenglonProducto <> 0 And FlexProduct.TextMatrix(mRenglonProducto, 6) = "" Then
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
      For mj = 2 To FlexIngreso.Rows - 1
         If FlexIngreso.TextMatrix(mj, 6) = FlexProduct.TextMatrix(mRenglonProducto, 6) And FlexIngreso.TextMatrix(mj, 7) = FlexProduct.TextMatrix(mRenglonProducto, 7) Then
            mRet = False
            mMensajeError = "El Producto y la Ubicación elegidos ya han sido seleccionados"
            mj = 999
         End If
      Next
   End If
      
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   fValidaIngreso = mRet
End Function
