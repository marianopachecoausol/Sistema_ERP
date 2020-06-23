VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form MantElect09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Abastecimiento para O.T."
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
      TabIndex        =   20
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
      Left            =   6600
      TabIndex        =   19
      Top             =   8600
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   960
      TabIndex        =   17
      Top             =   4120
      Width           =   1575
   End
   Begin VB.CommandButton CommandProd 
      Height          =   375
      Index           =   0
      Left            =   2760
      Picture         =   "MantElect09.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4080
      Width           =   330
   End
   Begin VB.CommandButton CommandProd 
      Height          =   375
      Index           =   1
      Left            =   3240
      Picture         =   "MantElect09.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
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
      Left            =   12000
      TabIndex        =   14
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
      Left            =   8640
      TabIndex        =   13
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
      Left            =   7080
      MaxLength       =   9
      TabIndex        =   12
      Top             =   60
      Width           =   1215
   End
   Begin VB.Frame Frame11 
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
      Height          =   3780
      Left            =   120
      TabIndex        =   9
      Top             =   4600
      Width           =   16680
      Begin MSFlexGridLib.MSFlexGrid FlexEgreso 
         Height          =   3255
         Left            =   120
         TabIndex        =   10
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
      TabIndex        =   2
      Top             =   600
      Width           =   16680
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   315
         Left            =   15000
         TabIndex        =   5
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   7560
         TabIndex        =   4
         Top             =   420
         Width           =   7335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   420
         Width           =   3735
      End
      Begin MSFlexGridLib.MSFlexGrid FlexProduct 
         Height          =   2295
         Left            =   120
         TabIndex        =   6
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
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Retirar de:"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.ComboBox Combo3 
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
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   75
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
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
      Left            =   5280
      TabIndex        =   11
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "O.T.  -  Fecha:"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "MantElect09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mObj As New clMantElect
Dim mObjInven As New clInven
Dim mRec As New ADODB.Recordset

Dim mRenglonProducto As Integer
Dim mRenglonEgreso As Integer
Dim mCodProducto As String

Dim mCodTipoVale As String
Dim mOTtieneValeAsociado As Boolean


Private Sub Combo3_Click()
   Dim IdOT As Integer
   Dim mi As Integer
   Dim mResultado As Boolean
   
   For mi = FlexProduct.Rows To 3 Step -1
      FlexProduct.RemoveItem mi
   Next
   mRenglonProducto = 0
   
   IdOT = CInt(Left(Combo3.Text, 10))

   Set mRec = mObj.oEjecutarSelect("SELECT IdOT, NroVale,CodTipoVale FROM OT_Abastecimiento_H O Where IdOT = " & IdOT & "; ")
   
   If Not mRec.EOF Then
      mOTtieneValeAsociado = True
      Text3.Text = mRec!NroVale
      Text3.Enabled = False
      
      If mRec!CodTipoVale = "C" Then
         Option1.Value = True
         Option2.Value = False
      Else
         Option1.Value = False
         Option2.Value = True
      End If
      Option1.Enabled = False
      Option2.Enabled = False
      
      Command3(0).Enabled = False
      CommandProd(0).Enabled = False
      CommandProd(1).Enabled = False
   Else
      mOTtieneValeAsociado = False
      Text3.Enabled = True
      Text3.Text = ""
      Option1.Value = False
      Option2.Value = False
      Option1.Enabled = True
      Option2.Enabled = True
      Command3(0).Enabled = True
      CommandProd(0).Enabled = True
      CommandProd(1).Enabled = True
   End If
   mRec.Close
   
   For mi = FlexEgreso.Rows To 3 Step -1
      FlexEgreso.RemoveItem mi
   Next
   mRenglonEgreso = 0
   
   If mOTtieneValeAsociado Then
      Set mRec = mObj.oEjecutarSelect("SELECT P.CodigoSap, AUX.CodProducto, AUX.CodUbicacion, P.Descripcion AS Producto, " & _
                                      "U.Descripcion AS Ubicacion, Med.Descripcion AS UnidadMedida,  AUX.Cantidad AS Cantidad  " & _
                                      "FROM (  " & _
                                      "      SELECT CodProducto, CodUbicacion, SUM(Cantidad) AS Cantidad  " & _
                                      "      From  " & _
                                      "      (  " & _
                                      "        SELECT MV.CodProducto, MV.CodUbicacion,  " & _
                                      "        CASE WHEN CodTipoMovimiento = 'I' THEN MV.Cantidad*(-1) ELSE MV.Cantidad END AS Cantidad  " & _
                                      "        From  " & _
                                      "          MantElect.OT_Abastecimiento_Det OM  " & _
                                      "        Inner Join  " & _
                                      "          Inventario.Movimientos2 MV ON OM.IdMov = MV.IdMov  " & _
                                      "        Where IdOT = " & IdOT & " " & _
                                      "      ) AS OT  " & _
                                      "      GROUP BY OT.CodProducto, OT.CodUbicacion  " & _
                                      "    ) AS AUX  " & _
                                      "INNER JOIN Inventario.Producto P ON AUX.CodProducto = P.Codigo  " & _
                                      "INNER JOIN Inventario.Ubicaciones U ON  AUX.CodUbicacion = U.Codigo  " & _
                                      "INNER JOIN Inventario.UnidadMedida Med ON P.CodUnidadMedida = Med.Codigo;  ")
         If Not mRec.EOF Then
            mi = 1
            Do While Not mRec.EOF
               mi = mi + 1
               
               With FlexEgreso
                  .AddItem ""
                  .TextMatrix(mi, 1) = mRec!CodigoSap
                  .TextMatrix(mi, 2) = mRec!Producto
                  .TextMatrix(mi, 3) = mRec!Ubicacion
                  .TextMatrix(mi, 4) = mRec!Cantidad
                  .TextMatrix(mi, 5) = mRec!UnidadMedida
                  .TextMatrix(mi, 6) = mRec!CodProducto
                  .TextMatrix(mi, 7) = mRec!CodUbicacion
               End With
         
               mRec.MoveNext
            Loop
         End If
         mRec.Close
   End If
End Sub

Private Sub Command1_Click()
   Dim mi As Integer
   Dim mj As Integer
   
   If Combo3.ListIndex <> -1 And Combo2.ListIndex <> -1 Then
      sMsgEspere Me, "Buscando productos...", True
      mRenglonProducto = 0
      
      'Elimino los registros (de la consulta anterior) de la grilla superior
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
      For mi = 2 To FlexProduct.Rows - 1
         For mj = 2 To FlexEgreso.Rows - 1
            If FlexProduct.TextMatrix(mi, 6) = FlexEgreso.TextMatrix(mj, 6) And FlexProduct.TextMatrix(mi, 7) = FlexEgreso.TextMatrix(mj, 7) Then
               FlexProduct.TextMatrix(mi, 4) = CDbl(Replace(Trim(FlexProduct.TextMatrix(mi, 4)), ".", ",")) - CDbl(Replace(Trim(FlexEgreso.TextMatrix(mj, 4)), ".", ","))
               mj = 999
            End If
         Next
      Next
      sMsgEspere Me, "", False
   Else
      MsgBox "Debe seleccionar alguna Orden de Trabajo y Lugar a retirar para poder buscar Productos.", vbExclamation, "Buscar Productos"
   End If
End Sub

Private Sub Command3_Click(Index As Integer)
If Index = 0 Then
   Dim vEgresosCodProducto() As String
   Dim vEgresosCodUbicacion() As String
   Dim vEgresosCantidad() As Double
   Dim vEgresosCantidadBD() As Double
   Dim cantEgresos As Integer
   Dim mj As Integer
   Dim mCodTipoVale As String
   Dim IdOT As Integer
   Dim CodUbicacion As String
   Dim mResultado As Boolean
   
   IdOT = CInt(Left(Combo3.Text, 10))
   CodUbicacion = Right(Trim(Combo2.Text), 4)

   If fValidaConfirmarAbastecimientoOT() Then
      cantEgresos = FlexEgreso.Rows - 2
   
   'pvMat_CantidadBD(mj - 2) = CDbl(Replace(FlexProduct.TextMatrix(mj, 8), ".", ","))
   
      ReDim vEgresosCodProducto(0 To cantEgresos - 1) As String
      ReDim vEgresosCodUbicacion(0 To cantEgresos - 1) As String
      ReDim vEgresosCantidad(0 To cantEgresos - 1) As Double
      ReDim vEgresosCantidadBD(0 To cantEgresos - 1) As Double
      
      For mj = 2 To FlexEgreso.Rows - 1
         vEgresosCodProducto(mj - 2) = FlexEgreso.TextMatrix(mj, 6)
         vEgresosCodUbicacion(mj - 2) = FlexEgreso.TextMatrix(mj, 7)
         vEgresosCantidad(mj - 2) = CDbl(Replace(FlexEgreso.TextMatrix(mj, 4), ".", ","))
         vEgresosCantidadBD(mj - 2) = CDbl(Replace(FlexEgreso.TextMatrix(mj, 8), ".", ","))
      Next
      
      If Option1.Value Then
         mCodTipoVale = "C"
      Else
         mCodTipoVale = "M"
      End If
      
      mResultado = True
      mObj.xInsOT_Abastecimiento vEgresosCodProducto(), vEgresosCodUbicacion(), vEgresosCantidad(), vEgresosCantidadBD(), Trim(Text3.Text), mCodTipoVale, IdOT, CodUbicacion, Trim(Right(MDI.mUser, 15)), mResultado

      If mResultado Then
         
         mOTtieneValeAsociado = True
         Command3(0).Enabled = False
         CommandProd(0).Enabled = False
         CommandProd(1).Enabled = False
         Text3.Enabled = False
         Option1.Enabled = False
         Option2.Enabled = False
         MsgBox "El consumo para esa O.T. se ha realizado exitosamente", vbInformation, "OT-Abastecimiento"
         
         'limpioFormulario
         'VerificaStockMin vEgresosCodProducto(), mCodBodega
      End If
   End If
Else
   Unload Me
End If

End Sub


Private Function fValidaConfirmarAbastecimientoOT() As Boolean
  'Validado: 'Que se haya completado el campo Numero Vale.
  'Validado: 'Que el Numero de Vale sea un valor entero.
  'Validado:  'Que esten chequeados alguna de las dos radio button (Option1 u Option2)
  'Validado: 'Que en la grilla inferior "Egresos" exista al menos un registro.
  'Validado: 'Que no exista el Registro (Nro Vale, TipoVale ) en la tabla OT_Abastecimiento_H
 
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
  
      Set mRec1 = mObj.oEjecutarSelect("SELECT NroVale,CodTipoVale FROM OT_Abastecimiento_H WHERE NroVale = " & Trim(Text3.Text) & " and CodTipoVale = '" & mCodTipoVale & "'; ")
      If Not mRec1.EOF Then
         mRet = False
         mMensajeError = "Ya se han realizado consumos para ese Número y Tipo de Vale !!!"
      End If
      mRec1.Close
   End If

   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   
   fValidaConfirmarAbastecimientoOT = mRet
End Function


Private Sub CommandProd_Click(Index As Integer)
Dim iStock As Double
   Dim mi As Integer
   Dim mRec1 As New ADODB.Recordset
   
   If Index = 0 Then
      If fValidaEgreso() Then
         FlexEgreso.AddItem vbTab & FlexProduct.TextMatrix(mRenglonProducto, 1) & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 2) & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 3) & vbTab & Text2.Text & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 5) & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 6) & vbTab & FlexProduct.TextMatrix(mRenglonProducto, 7) & vbTab & 0
         FlexProduct.TextMatrix(mRenglonProducto, 4) = Format(CDbl(Replace(Trim(FlexProduct.TextMatrix(mRenglonProducto, 4)), ".", ",")) - CDbl(Replace(Trim(Text2.Text), ".", ",")), "00.00")
         Text2.Text = ""
         Text2.SetFocus
      End If
   Else
      For mi = 2 To FlexProduct.Rows - 1
      
         If FlexProduct.TextMatrix(mi, 6) = FlexEgreso.TextMatrix(mRenglonEgreso, 6) And FlexProduct.TextMatrix(mi, 7) = FlexEgreso.TextMatrix(mRenglonEgreso, 7) Then
            Set mRec1 = mObj.oEjecutarSelect("SELECT Fecha,CodProducto,CodUbicacion, IFNULL(Stock,0) As Stock " & _
                                       " FROM Inventario.Movimientos2 M " & _
                                       " WHERE CodProducto  = '" & FlexEgreso.TextMatrix(mRenglonEgreso, 6) & "' and CodUbicacion = '" & FlexEgreso.TextMatrix(mRenglonEgreso, 7) & "'" & _
                                       " AND Fecha = (SELECT Max(Fecha) FROM Inventario.Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion)")
      
            If Not mRec1.EOF Then
               iStock = mRec1!Stock
            Else
               iStock = 0
            End If
            mRec1.Close
            
            FlexProduct.TextMatrix(mi, 4) = iStock
           
            mi = 9999
         End If
      Next
      
      If FlexEgreso.Rows > 2 And mRenglonEgreso > 1 Then
         FlexEgreso.RemoveItem (mRenglonEgreso)
      End If
      
      mRenglonEgreso = 0
   End If

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
   
   Command3(0).Enabled = False
   CommandProd(0).Enabled = False
   CommandProd(1).Enabled = False
   
   
   Set mRec = mObj.oEjecutarSelect("SELECT CONVERT( CONCAT(LPAD(IdOT,10,'0'),' - ',Date_Format(Fecha,'%d/%m/%Y')),char(23)) AS OT_Fecha, O.Fecha " & _
                           " FROM MantElect.OT_H O " & _
                           " ORDER BY IdOT DESC; ")

   Do While Not mRec.EOF
      Combo3.AddItem mRec!OT_Fecha
      mRec.MoveNext
   Loop
   mRec.Close
   
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

   With FlexEgreso
      .ColWidth(0) = 200
      .ColWidth(1) = 950
      .ColWidth(2) = 9700
      .ColWidth(3) = 2100
      .ColWidth(4) = 1200
      .ColWidth(5) = 1900
    
      .ColWidth(6) = 0
      .ColWidth(7) = 0
      .ColWidth(8) = 0
      '.ColWidth(9) = 1000
      '.ColWidth(10) = 1000
   
      .TextMatrix(0, 1) = "Cód.Sap"
      .TextMatrix(0, 2) = "Producto"
      .TextMatrix(0, 3) = "Ubicación"
      .TextMatrix(0, 4) = "Cantidad"
      .TextMatrix(0, 5) = "Unid.Medida"
      .TextMatrix(0, 6) = "Cód. Producto"
      .TextMatrix(0, 7) = "Cód. Ubicacion"
      .TextMatrix(0, 8) = "CantidadBD"
      '.TextMatrix(0, 9) = "StockActual"
      '.TextMatrix(0, 10) = "YaEstaEnOT"

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
      Set mRec1 = mObj.oEjecutarSelect("SELECT NroVale,CodTipoVale FROM OT_Abastecimiento_H WHERE NroVale = " & Trim(Text3.Text) & " and CodTipoVale = '" & mCodTipoVale & "'; ")
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
                                       " FROM Inventario.Movimientos2 M " & _
                                       " WHERE CodProducto  = '" & FlexProduct.TextMatrix(mRenglonProducto, 6) & "' and CodUbicacion = '" & FlexProduct.TextMatrix(mRenglonProducto, 7) & "'" & _
                                       " AND Fecha = (SELECT Max(Fecha) FROM Inventario.Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion)")
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
