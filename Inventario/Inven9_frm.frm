VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Inven9_frm 
   Caption         =   "Stock Mínimo"
   ClientHeight    =   10290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15720
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   15720
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   9220
      TabIndex        =   11
      Top             =   9600
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stock Mínimo"
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
      Height          =   3735
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   15495
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   278
         Left            =   13800
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   3360
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid FlexStMin 
         Height          =   2895
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   5106
         _Version        =   327680
         Cols            =   8
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   15495
      Begin VB.CommandButton Command2 
         Caption         =   "Buscar "
         Height          =   375
         Left            =   10320
         TabIndex        =   5
         Top             =   520
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
         Left            =   1920
         TabIndex        =   4
         Top             =   520
         Width           =   8175
      End
      Begin MSFlexGridLib.MSFlexGrid FlexProduct 
         Height          =   3135
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   5530
         _Version        =   327680
         Cols            =   5
      End
      Begin VB.Label Label1 
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
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Index           =   0
      Left            =   5420
      TabIndex        =   0
      Top             =   9600
      Width           =   1095
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   5040
      Width           =   9495
   End
   Begin VB.Label Label2 
      Caption         =   "Stock Mínimo para producto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   3495
   End
End
Attribute VB_Name = "Inven9_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mObj As New clInven
Dim mRec As ADODB.Recordset
Dim mRenglonProducto As Integer
Dim mCodProducto As String

Dim filaAnt As Integer
Dim columnAnt As Integer

Private Sub Command1_Click(Index As Integer)
   Dim mi As Integer
   Dim CodProducto As String
   Dim CodBodega As String
   Dim stockMin As String
   Dim stockMinInicial As String
   
   If Index = 0 Then
      'En este caso 6 es la columna que seria editable
      If columnAnt = 6 Then
         FlexStMin.TextMatrix(filaAnt, columnAnt) = Text2.Text
      End If
   
      Text2.Visible = False
      If fValidaStockMin() Then
         For mi = 2 To FlexStMin.Rows - 1
            
            CodBodega = FlexStMin.TextMatrix(mi, 4)
            CodProducto = FlexStMin.TextMatrix(mi, 5)
            stockMin = FlexStMin.TextMatrix(mi, 6)
            stockMinInicial = FlexStMin.TextMatrix(mi, 7)
            
            'Si hice alguna modificion entonces actualizo la BD.
            If CDbl(Replace(Trim(stockMin), ".", ",")) <> CDbl(Replace(Trim(stockMinInicial), ".", ",")) Then
               'Si CodProducto = "" and StockMinino <> 0 => Inserto en BD
               If CodProducto = "" And CDbl(Replace(Trim(stockMin), ".", ",")) <> 0 Then
                  mObj.xInsStockMinimo mCodProducto, CodBodega, stockMin
                  FlexStMin.TextMatrix(mi, 5) = mCodProducto
               End If
               
               'Si CodProducto <> "" and StockMinino <> 0 => Actualizo  BD
               If CodProducto <> "" And CDbl(Replace(Trim(stockMin), ".", ",")) <> 0 Then
                  mObj.xUpdStockMinimo mCodProducto, CodBodega, stockMin
               End If
               
               'Si CodProducto <> "" and StockMinino = 0 => Elimino   BD
               If CodProducto <> "" And CDbl(Replace(Trim(stockMin), ".", ",")) = 0 Then
                  mObj.xDelStockMinimo mCodProducto, CodBodega
                  FlexStMin.TextMatrix(mi, 5) = ""
               End If
            End If
            
            
            mObj.xDelStockMinimo_Notificaciones CodProducto, CodBodega

         Next
         MsgBox "Se han realizado las modificaciones exitosamente", vbInformation, "Stock Mínimo"
      End If
   Else
      Unload Me
   End If
End Sub

Private Function fValidaStockMin() As Boolean
   Dim mi As Integer
   Dim mRet As Boolean
   Dim mMensajeError As String
   Dim posInstr As Integer
   Dim qtyDecimales As Integer

   mRet = True
      
   If FlexStMin.Rows <= 2 Then
      mRet = False
      mMensajeError = "Al menos debe existir un producto en la grilla inferior"
   End If
      
   For mi = 2 To FlexStMin.Rows - 1
      If mRet Then
         If Not IsNumeric(Replace(FlexStMin.TextMatrix(mi, 6), ".", ",")) Then
            mRet = False
            mMensajeError = "Se ha cargado un valor incorrecto para la bodega: '" & FlexStMin.TextMatrix(mi, 2) & "'"
         End If
      End If
   Next
   
   For mi = 2 To FlexStMin.Rows - 1
      'Valido que no supere los 2 digitos decimales
      If mRet Then
         posInstr = InStr(1, Replace(FlexStMin.TextMatrix(mi, 6), ".", ","), ",")
      
         If posInstr <> 0 Then
            qtyDecimales = Len(Right(Trim(FlexStMin.TextMatrix(mi, 6)), Len(Trim(FlexStMin.TextMatrix(mi, 6))) - posInstr))
         End If
   
         If qtyDecimales > 2 Then
            mRet = False
            mMensajeError = "El Stock Mínimo para la Bodega: '" & FlexStMin.TextMatrix(mi, 2) & "' tiene más de dos dígitos decimales."
         End If
      End If
   Next
   
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   fValidaStockMin = mRet

End Function


Private Sub Command2_Click()
   
   Dim mi As Integer
   
   mRenglonProducto = 0
   
   For mi = FlexProduct.Rows To 3 Step -1
      FlexProduct.RemoveItem mi
   Next
   
   Set mRec = mObj.oEjecutarSelect("SELECT P.Codigo, P.Descripcion, P.CodigoSap, U.Descripcion as UnidadMedida " & _
     " From " & _
     " Producto P  " & _
     " Inner Join  " & _
     " UnidadMedida U ON P.CodUnidadMedida = U.Codigo  " & _
     " where P.Descripcion like '%" & Text1.Text & "%'  " & _
     " and P.Fecha_Baja is null " & _
     " ORDER BY P.Descripcion; ")
         
   If Not mRec.EOF Then
      mi = 1
      Do While Not mRec.EOF
         mi = mi + 1
         
         FlexProduct.AddItem ""
         FlexProduct.TextMatrix(mi, 1) = mRec!descripcion
         FlexProduct.TextMatrix(mi, 2) = mRec!UnidadMedida
         FlexProduct.TextMatrix(mi, 3) = mRec!CodigoSap
         FlexProduct.TextMatrix(mi, 4) = mRec!Codigo
       
         mRec.MoveNext
      Loop
   End If
   mRec.Close
   
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
         Label2.Visible = True
         Label3.Visible = True
         Label3.Caption = FlexProduct.TextMatrix(mRenglonProducto, 1)
         mCodProducto = FlexProduct.TextMatrix(mRenglonProducto, 4)
      
      
         Text2.Visible = False
      
         For mi = FlexStMin.Rows To 3 Step -1
            FlexStMin.RemoveItem mi
         Next
        
        Set mRec = mObj.oEjecutarSelect("SELECT B.CodAlmacen, A.Descripcion As Almacen, " & _
        " B.Codigo As CodBodega, B.Descripcion As Bodega, " & _
        " IFNULL(S.CodProducto,'') AS CodProducto, " & _
        " IFNULL(S.Stock_Min, 0) As Stock_Min " & _
        " From " & _
        " Bodegas B  " & _
        " Left Join  " & _
        " StocksMinimos S ON B.Codigo = S.CodBodega and S.CodProducto = '" & mCodProducto & "' " & _
        " Inner Join  " & _
        " Almacenes A ON A.Codigo = B.CodAlmacen " & _
        " WHERE B.Codigo IN (SELECT CodBodega FROM Usuario_AccesoBodega WHERE CodUsuario = '" & Trim(Right(MDI.mUser, 15)) & "')  ")
   
        '" and .Fecha_Baja is null " & _
        '" ORDER BY P.Descripcion; ")
         
         If Not mRec.EOF Then
            mi = 1
            Do While Not mRec.EOF
               mi = mi + 1
      
               FlexStMin.AddItem ""
               FlexStMin.TextMatrix(mi, 1) = mRec!Almacen
               FlexStMin.TextMatrix(mi, 2) = mRec!Bodega
               FlexStMin.TextMatrix(mi, 3) = mRec!CodAlmacen
               FlexStMin.TextMatrix(mi, 4) = mRec!CodBodega
               FlexStMin.TextMatrix(mi, 5) = mRec!CodProducto
               FlexStMin.TextMatrix(mi, 6) = mRec!Stock_Min
               FlexStMin.TextMatrix(mi, 7) = mRec!Stock_Min
      
               mRec.MoveNext
            Loop
            'FlexProduct.RemoveItem 1
         End If
         mRec.Close
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
   Dim mi As Integer
   
   Inven9_frm.Width = 15840
   Inven9_frm.Height = 10800
   sAlinearForm Me
   
   
   filaAnt = 0
   columnAnt = 0
   Text2.Visible = False
   
   Label2.Visible = False
   Label3.Visible = False
   
   FlexProduct.ColWidth(0) = 200
   FlexProduct.ColWidth(1) = 10000
   FlexProduct.ColWidth(2) = 2000
   FlexProduct.ColWidth(3) = 0
   FlexProduct.ColWidth(4) = 2000
   
   FlexProduct.TextMatrix(0, 1) = "Producto"
   FlexProduct.TextMatrix(0, 2) = "Unidad de Medida"
   FlexProduct.TextMatrix(0, 3) = "Código Sap"
   FlexProduct.TextMatrix(0, 4) = "Código"
   
   FlexProduct.RowHeight(1) = 0
   
   FlexStMin.ColWidth(0) = 200
   FlexStMin.ColWidth(1) = 6940
   FlexStMin.ColWidth(2) = 6000
   FlexStMin.ColWidth(3) = 0
   FlexStMin.ColWidth(4) = 0
   FlexStMin.ColWidth(5) = 0
   FlexStMin.ColWidth(6) = 1700
   FlexStMin.ColWidth(7) = 0
   
   FlexStMin.TextMatrix(0, 1) = "Almacén"
   FlexStMin.TextMatrix(0, 2) = "Bodega"
   FlexStMin.TextMatrix(0, 3) = "CodAlmacen"
   FlexStMin.TextMatrix(0, 4) = "CodBodega"
   FlexStMin.TextMatrix(0, 5) = "CodProducto"
   FlexStMin.TextMatrix(0, 6) = "Stock Mínimo"
   FlexStMin.TextMatrix(0, 7) = "Stock Mínimo Inicial"

   FlexStMin.RowHeight(1) = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mObj = Nothing
   Set mRec = Nothing
   
   ShowMenu 12, True, False
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
      If KeyAscii <> 46 Then
         KeyAscii = fNumeroKeyPress(KeyAscii)
      End If
      
      If KeyAscii = 13 Then
         FlexStMin.TextMatrix(filaAnt, columnAnt) = Text2.Text
         Text2.Visible = False
         FlexStMin.ScrollBars = flexScrollBarVertical
      End If
End Sub

Private Sub Text2_LostFocus()
   If FlexStMin.Col <> columnAnt Or FlexStMin.Row <> filaAnt Then
      
      'En este caso 6 es la columna que seria editable
      If columnAnt = 6 Then
         FlexStMin.TextMatrix(filaAnt, columnAnt) = Text2.Text
         'FlexStMin.ScrollBars = flexScrollBarNone
      End If
   End If
End Sub


Private Sub FlexStMin_Click()
   
   'En este caso 6 es la columna que seria editable
   If FlexStMin.Col = 6 And FlexStMin.Row <> 1 Then
      Text2.Text = FlexStMin.Text
      Text2.Width = FlexStMin.ColWidth(FlexStMin.Col)
      Text2.Left = FlexStMin.ColPos(FlexStMin.Col) + FlexStMin.Left + 30 'el valor treina termina de acomodar el textbox en la celda
      Text2.Top = FlexStMin.Top + FlexStMin.RowPos(FlexStMin.Row)
      Text2.Visible = True
      Text2.SetFocus
      FlexStMin.ScrollBars = flexScrollBarNone
   Else
      Text2.Visible = False
      FlexStMin.ScrollBars = flexScrollBarVertical
      
   End If

   filaAnt = FlexStMin.Row
   columnAnt = FlexStMin.Col
End Sub

