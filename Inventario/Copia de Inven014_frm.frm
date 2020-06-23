VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Inven014_frmGrande 
   Caption         =   "Agregar Items a Orden de Compra"
   ClientHeight    =   13440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   21150
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   13440
   ScaleWidth      =   21150
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1(uno)"
      Height          =   11375
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   20895
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
         TabIndex        =   20
         Top             =   10800
         Width           =   2175
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
         Left            =   6840
         TabIndex        =   19
         Top             =   10800
         Width           =   2175
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   360
         Top             =   -240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483633
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Inven014_frm.frx":0000
               Key             =   "Accept"
               Object.Tag             =   "Accept"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Inven014_frm.frx":059A
               Key             =   "Add"
               Object.Tag             =   "Add"
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame4 
         Caption         =   "Nuevos Items"
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
         Left            =   0
         TabIndex        =   15
         Top             =   6180
         Width           =   20895
         Begin MSFlexGridLib.MSFlexGrid FlexIngreso 
            Height          =   3615
            Left            =   240
            TabIndex        =   16
            Top             =   600
            Width           =   20415
            _ExtentX        =   36010
            _ExtentY        =   6376
            _Version        =   327680
            Cols            =   8
         End
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Index           =   0
         Left            =   3000
         Picture         =   "Inven014_frm.frx":0B34
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5580
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Index           =   1
         Left            =   3500
         Picture         =   "Inven014_frm.frx":0E3E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5580
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
         TabIndex        =   12
         Top             =   5580
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Caption         =   "Seleccion del Producto"
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
         Left            =   0
         TabIndex        =   6
         Top             =   120
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
            TabIndex        =   9
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
            TabIndex        =   8
            Top             =   600
            Width           =   10455
         End
         Begin MSFlexGridLib.MSFlexGrid FlexProduct 
            Height          =   3735
            Left            =   240
            TabIndex        =   10
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
            TabIndex        =   7
            Top             =   600
            Width           =   1455
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
         TabIndex        =   11
         Top             =   5640
         UseMnemonic     =   0   'False
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1(cero)"
      Height          =   11175
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   20895
      Begin VB.Frame Frame5 
         Caption         =   "Recepciones de la Orden de Compra"
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
         Height          =   10575
         Left            =   0
         TabIndex        =   17
         Top             =   120
         Width           =   20895
         Begin MSFlexGridLib.MSFlexGrid FlexRecepcionesEfectuadas 
            Height          =   5655
            Left            =   240
            TabIndex        =   18
            Top             =   720
            Width           =   20415
            _ExtentX        =   36010
            _ExtentY        =   9975
            _Version        =   327680
            Cols            =   9
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   20895
      _ExtentX        =   36856
      _ExtentY        =   661
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Recepciones Efectuadas"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Agregar Items a Orden de Compra"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selección de Orden de Compra"
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
      Height          =   1305
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   20895
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
         TabIndex        =   2
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label1 
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
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Inven014_frmGrande"
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
Dim mCodBodega As String


Private Sub Combo1_Click()
   Dim sNroOC As String
   Dim mi As Integer
   Dim i As Integer
      
   sNroOC = Trim(Left(Combo1.Text, 10))
     
   For i = 1 To 10
    sNroOC = Replace(sNroOC, "_", "")
   Next
      
   'Visualizo el primer frame
   Frame1(0).Visible = True
   Frame1(1).Visible = False
       
   'Posiciono el tabstrip1 en la posicion 1 para que quede seleccionada la solapa
   TabStrip1.Tabs(1).Selected = True
       
   'Elimino los registros (de la consulta anterior) de la grilla superior
   For mi = FlexRecepcionesEfectuadas.Rows To 3 Step -1
      FlexRecepcionesEfectuadas.RemoveItem mi
   Next

   Set mRec = mObj.oEjecutarSelect(" SELECT Date_Format(M.Fecha,'%d-%m-%Y') AS FechaRecepcion, P.Descripcion AS Producto, " & _
                           " M.Cantidad, UB.Descripcion AS Ubicacion,U.Descripcion AS UnidadMedida, P.CodigoSap " & _
                           " FROM " & _
                           "  Ingresos_Det ID " & _
                           " INNER JOIN " & _
                           "  Movimientos2 M ON ID.IdMov = M.IdMov " & _
                           " INNER JOIN " & _
                           "  Producto P ON P.Codigo = M.CodProducto " & _
                           " INNER JOIN " & _
                           "  Ubicaciones UB ON UB.Codigo = M.CodUbicacion " & _
                           " INNER JOIN " & _
                           "   UnidadMedida U ON U.Codigo = P.CodUnidadMedida " & _
                           " WHERE NroOC = '" & sNroOC & "' " & _
                           " ORDER BY M.Fecha, P.Descripcion; ")

   'Cargo la Grilla del Panel de Recepciones Efectuadas
   If Not mRec.EOF Then
      mi = 1
      Do While Not mRec.EOF
         mi = mi + 1

         FlexRecepcionesEfectuadas.AddItem ""
         FlexRecepcionesEfectuadas.TextMatrix(mi, 1) = mRec!FechaRecepcion
         FlexRecepcionesEfectuadas.TextMatrix(mi, 2) = mRec!Producto
         FlexRecepcionesEfectuadas.TextMatrix(mi, 3) = mRec!Ubicacion
         FlexRecepcionesEfectuadas.TextMatrix(mi, 4) = mRec!Cantidad
         FlexRecepcionesEfectuadas.TextMatrix(mi, 5) = mRec!UnidadMedida
         FlexRecepcionesEfectuadas.TextMatrix(mi, 6) = mRec!CodigoSap
         'FlexRecepcionesEfectuadas.TextMatrix(mI, 7) = mRec!CodUbicacion

         mRec.MoveNext
      Loop
   End If
   mRec.Close
      
      
   'Busco el codigo de bodega de Orden de compra, para poder utilizarla en el boton Buscar de la grilla producto
   Set mRec = mObj.oEjecutarSelect(" SELECT * FROM Inventario.Ingresos_H I where NroOC = " & sNroOC & ";")
   mCodBodega = mRec!CodBodega
   mRec.Close
      
   'TODO(Realizado): Elimiar TextBox:"Contiene" del Segundo Frame Frame1(1)
   Text1.Text = ""
   Text2.Text = ""
   
   'TODO(Realizado): Elimiar Datos de la grilla Seleccion del Producto del Segundo Frame: Frame1(1)
   For mi = FlexProduct.Rows To 3 Step -1
      FlexProduct.RemoveItem mi
   Next
   
   mRenglonProducto = 0

   'TODO(Realizado): Elimiar Datos de la grilla Ingresos del Producto del Segundo Frame: Frame1(1)
   For mi = FlexIngreso.Rows To 3 Step -1
      FlexIngreso.RemoveItem mi
   Next
   mRenglonEgreso = 0
   
End Sub

Private Sub Command1_Click()
   Dim mi As Integer
   Dim mJ As Integer
   
   mRenglonProducto = 0
   
   'Elimino los registros (de la consulta anterior) de la grilla superior
   For mi = FlexProduct.Rows To 3 Step -1
      FlexProduct.RemoveItem mi
   Next
   
   'Set mRec = mObj.getStockXBodegaConFiltroProducto(Left(Combo1.Text, 4), Text1.Text)
   'TODO(Realizado): Agregar la bodega que se va a obtener del label que muestre la bodega de esa OC, ese label podria estar oculto)
   Set mRec = mObj.getStockXBodegaConFiltroProducto(mCodBodega, Text1.Text)
   
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
      For mJ = 2 To FlexIngreso.Rows - 1
         If FlexProduct.TextMatrix(mi, 6) = FlexIngreso.TextMatrix(mJ, 6) And FlexProduct.TextMatrix(mi, 7) = FlexIngreso.TextMatrix(mJ, 7) Then
            FlexProduct.TextMatrix(mi, 3) = CDbl(Replace(Trim(FlexProduct.TextMatrix(mi, 3)), ".", ",")) - CDbl(Replace(Trim(FlexIngreso.TextMatrix(mJ, 3)), ".", ","))
            mJ = 999
         End If
      Next
   Next
End Sub

Private Sub Command2_Click(Index As Integer)
   
   Dim iStock As Double
   Dim mi As Integer
   Dim mRec1 As New ADODB.Recordset
   
   If Index = 0 Then
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
'         If FlexProduct.TextMatrix(mi, 6) = FlexIngreso.TextMatrix(mRenglonEgreso, 6) And FlexProduct.TextMatrix(mi, 7) = FlexIngreso.TextMatrix(mRenglonEgreso, 7) Then
'            Set mRec1 = mObj.oEjecutarSelect("SELECT Fecha,CodProducto,CodUbicacion, IFNULL(Stock,0) As Stock " & _
'                                       " FROM Movimientos2 M " & _
'                                       " WHERE CodProducto  = '" & FlexIngreso.TextMatrix(mRenglonEgreso, 6) & "' and CodUbicacion = '" & FlexIngreso.TextMatrix(mRenglonEgreso, 7) & "'" & _
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
      
      If FlexIngreso.Rows > 2 And mRenglonEgreso > 1 Then
         FlexIngreso.RemoveItem (mRenglonEgreso)
      End If
      
      mRenglonEgreso = 0
   End If

End Sub

Private Sub Command3_Click(Index As Integer)
   If Index = 0 Then
      Dim sNroOC As String
      Dim vIngresosCodProducto() As String
      Dim vIngresosCodUbicacion() As String
      Dim vIngresosCantidad() As Double
      Dim cantIngresos As Integer
      Dim mJ As Integer
      Dim mi As Integer
      Dim i As Integer
      Dim mResultado As Boolean
      
   
      If fValidaConfirmarIngreso Then
         sNroOC = Trim(Left(Combo1.Text, 10))
   
         For i = 1 To 10
          sNroOC = Replace(sNroOC, "_", "")
         Next
         
         cantIngresos = FlexIngreso.Rows - 2
      
         ReDim vIngresosCodProducto(0 To cantIngresos - 1) As String
         ReDim vIngresosCodUbicacion(0 To cantIngresos - 1) As String
         ReDim vIngresosCantidad(0 To cantIngresos - 1) As Double
         
         
         For mJ = 2 To FlexIngreso.Rows - 1
            vIngresosCodProducto(mJ - 2) = FlexIngreso.TextMatrix(mJ, 6)
            vIngresosCodUbicacion(mJ - 2) = FlexIngreso.TextMatrix(mJ, 7)
            vIngresosCantidad(mJ - 2) = CDbl(Replace(FlexIngreso.TextMatrix(mJ, 3), ".", ","))
         Next
         
         
         mResultado = True
         
         mObj.xInsAddItemsOC vIngresosCodProducto(), vIngresosCodUbicacion(), vIngresosCantidad(), sNroOC, Trim(Right(MDI.mUser, 15)), mResultado
         
         If mResultado Then
            MsgBox "Los nuevos items se han agregado exitosamente", vbInformation, "Nuevos Items"
            actualizaFlagStockMinimo vIngresosCodProducto(), mCodBodega
            
            Text1.Text = ""
            Text2.Text = ""
                        
           'Visualizo el primer frame
            Frame1(0).Visible = True
            Frame1(1).Visible = False
          
            'Posiciono el tabstrip1 en la posicion 1 para que quede seleccionada la solapa
            TabStrip1.Tabs(1).Selected = True
            
            'TODO(Realizado): Elimiar Datos de la grilla Seleccion del Producto del Segundo Frame: Frame1(1)
            For mi = FlexProduct.Rows To 3 Step -1
               FlexProduct.RemoveItem mi
            Next
   
            mRenglonProducto = 0

            'TODO(Realizado): Elimiar Datos de la grilla Ingresos del Producto del Segundo Frame: Frame1(1)
            For mi = FlexIngreso.Rows To 3 Step -1
               FlexIngreso.RemoveItem mi
            Next
            mRenglonEgreso = 0
            
            
            'Elimino los registros (de la consulta anterior) de la grilla superior
            For mi = FlexRecepcionesEfectuadas.Rows To 3 Step -1
               FlexRecepcionesEfectuadas.RemoveItem mi
            Next
      
            Set mRec = mObj.oEjecutarSelect(" SELECT Date_Format(M.Fecha,'%d-%m-%Y') AS FechaRecepcion, P.Descripcion AS Producto, " & _
                              " M.Cantidad, UB.Descripcion AS Ubicacion,U.Descripcion AS UnidadMedida, P.CodigoSap " & _
                              " FROM " & _
                              "  Ingresos_Det ID " & _
                              " INNER JOIN " & _
                              "  Movimientos2 M ON ID.IdMov = M.IdMov " & _
                              " INNER JOIN " & _
                              "  Producto P ON P.Codigo = M.CodProducto " & _
                              " INNER JOIN " & _
                              "  Ubicaciones UB ON UB.Codigo = M.CodUbicacion " & _
                              " INNER JOIN " & _
                              "   UnidadMedida U ON U.Codigo = P.CodUnidadMedida " & _
                              " WHERE NroOC = '" & sNroOC & "' " & _
                              " ORDER BY M.Fecha, P.Descripcion; ")
   
            'Cargo la Grilla del Panel de Recepciones Efectuadas
            If Not mRec.EOF Then
               mi = 1
               Do While Not mRec.EOF
                  mi = mi + 1
         
                  FlexRecepcionesEfectuadas.AddItem ""
                  FlexRecepcionesEfectuadas.TextMatrix(mi, 1) = mRec!FechaRecepcion
                  FlexRecepcionesEfectuadas.TextMatrix(mi, 2) = mRec!Producto
                  FlexRecepcionesEfectuadas.TextMatrix(mi, 3) = mRec!Ubicacion
                  FlexRecepcionesEfectuadas.TextMatrix(mi, 4) = mRec!Cantidad
                  FlexRecepcionesEfectuadas.TextMatrix(mi, 5) = mRec!UnidadMedida
                  FlexRecepcionesEfectuadas.TextMatrix(mi, 6) = mRec!CodigoSap
                  'FlexRecepcionesEfectuadas.TextMatrix(mI, 7) = mRec!CodUbicacion
         
                  mRec.MoveNext
               Loop
            End If
            mRec.Close
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





Private Sub flexIngreso_Click()
   Dim mi As Integer
   
   If FlexIngreso.MouseRow > 0 Then
   
      If mRenglonEgreso <> 0 Then
         If FlexIngreso.Rows > mRenglonEgreso Then
            FlexIngreso.Row = mRenglonEgreso
            For mi = 1 To FlexIngreso.Cols - 1
               FlexIngreso.Col = mi
               FlexIngreso.CellBackColor = vbWhite
            Next
         End If
      End If
      
      mRenglonEgreso = FlexIngreso.MouseRow
   
      FlexIngreso.Row = mRenglonEgreso
      For mi = 1 To FlexIngreso.Cols - 1
         FlexIngreso.Col = mi
         FlexIngreso.CellBackColor = &H80000003
      Next
      
      If mRenglonEgreso > 1 Then
          mCodProducto = FlexIngreso.TextMatrix(mRenglonEgreso, 4)
      End If
   Else
      FlexIngreso.Row = mRenglonEgreso
      For mi = 1 To FlexProduct.Cols - 1
         FlexIngreso.Col = mi
         FlexIngreso.CellBackColor = vbWhite
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
   Inven014_frm.Width = 21270
   Inven014_frm.Height = 13950
   sAlinearForm Me
   mCodBodega = ""
   
   Set mRec = mObj.oEjecutarSelect(" SELECT CONVERT( LPAD(CONCAT (NroOC, ' - ', Date_Format(Fecha,'%d/%m/%Y')),23,'_'),CHAR(23)) as OC_and_Fecha " & _
                           " FROM Ingresos_H I " & _
                           " ORDER BY Fecha; ")
   
   Do While Not mRec.EOF
      Combo1.AddItem mRec!OC_and_Fecha
      mRec.MoveNext
   Loop
   mRec.Close
   
   
   Frame1(0).Visible = True
   Frame1(1).Visible = False
   
   FlexRecepcionesEfectuadas.ColWidth(0) = 200
   
   FlexRecepcionesEfectuadas.ColWidth(1) = 1250
   
   FlexRecepcionesEfectuadas.ColWidth(2) = 9450
   FlexRecepcionesEfectuadas.ColWidth(3) = 4500
   FlexRecepcionesEfectuadas.ColWidth(4) = 1500
   FlexRecepcionesEfectuadas.ColWidth(5) = 1900
   FlexRecepcionesEfectuadas.ColWidth(6) = 1250
   FlexRecepcionesEfectuadas.ColWidth(7) = 0
   FlexRecepcionesEfectuadas.ColWidth(8) = 0
   
   FlexRecepcionesEfectuadas.TextMatrix(0, 1) = "Fecha Recep."
   FlexRecepcionesEfectuadas.TextMatrix(0, 2) = "Producto"
   FlexRecepcionesEfectuadas.TextMatrix(0, 3) = "Ubicación"
   FlexRecepcionesEfectuadas.TextMatrix(0, 4) = "Cantidad"
   FlexRecepcionesEfectuadas.TextMatrix(0, 5) = "Unid.Medida"
   FlexRecepcionesEfectuadas.TextMatrix(0, 6) = "Cód.Sap"
   FlexRecepcionesEfectuadas.TextMatrix(0, 7) = "Cód. Producto"
   FlexRecepcionesEfectuadas.TextMatrix(0, 8) = "Cód. Ubicacion"

   FlexRecepcionesEfectuadas.RowHeight(1) = 0
   
   
   
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
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mObj = Nothing
   Set mRec = Nothing
   ShowMenu 12, True, False
End Sub


Private Sub TabStrip1_Click()

   Dim i As Integer
   Dim j As Integer
    
    i = TabStrip1.SelectedItem.Index
  
   
   For j = 1 To TabStrip1.Tabs.Count
      If j = i Then
         Frame1(j - 1).Visible = True
         
      Else
         Frame1(j - 1).Visible = False
      End If
   Next

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


Private Sub Text2_KeyPress(KeyAscii As Integer)
      If KeyAscii <> 46 Then
         KeyAscii = fNumeroKeyPress(KeyAscii)
      End If
End Sub


Private Function fValidaConfirmarIngreso() As Boolean
   Dim mRet As Boolean
   Dim mMensajeError As String
   Dim mRec1 As New ADODB.Recordset
   
   mRet = True
   
   If Combo1.ListIndex = -1 Then
      mRet = False
      mMensajeError = "Debe seleccionar una Orden de Compra"
   End If
   
     
   If mRet Then
      If FlexIngreso.Rows <= 2 Then
         mRet = False
         mMensajeError = "Al menos debe existir un registro en la Grilla: 'Nuevos Items'"
      End If
   End If
   
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   
   fValidaConfirmarIngreso = mRet
End Function







