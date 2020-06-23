VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Inven5_frm 
   Caption         =   "Movimientos de Inventario - EGRESOS"
   ClientHeight    =   13920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   26550
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   13920
   ScaleWidth      =   26550
   Begin VB.CommandButton Command2 
      Caption         =   "Volver"
      Height          =   495
      Index           =   1
      Left            =   19200
      TabIndex        =   12
      Top             =   13080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aceptar"
      Height          =   495
      Index           =   0
      Left            =   6240
      TabIndex        =   11
      Top             =   13080
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Egresos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   9
      Top             =   8280
      Width           =   26295
      Begin MSFlexGridLib.MSFlexGrid FlexEgreso 
         Height          =   3615
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   25875
         _ExtentX        =   45641
         _ExtentY        =   6376
         _Version        =   327680
         Cols            =   7
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   10200
      MaxLength       =   60
      TabIndex        =   7
      Top             =   7560
      Width           =   7575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   7560
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   18240
      Picture         =   "Inven5_frm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Productos Disponibles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   26295
      Begin MSFlexGridLib.MSFlexGrid FlexProd 
         Height          =   6495
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   25875
         _ExtentX        =   45641
         _ExtentY        =   11456
         _Version        =   327680
         Cols            =   7
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   8
      Top             =   7560
      UseMnemonic     =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Motivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   7560
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   7560
      UseMnemonic     =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "Inven5_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mObj As New clInven
Dim mRec As New ADODB.Recordset
Dim mRenglon As Integer

Private Type Movimiento
    CodProducto As String
    Cantidad As Double
End Type

Dim vMovimientos() As Movimiento
Dim contDim As Integer


Private Sub Inicio()
   Dim mI As Integer
   
   Me.Width = 26670
   Me.Height = 14430
   Me.Top = 100
   Me.Left = (MDI.Width - Me.Width) / 2
   
   contDim = 0
   
   '--------------------------------------------------GRILLA PRODUCTOS DISPONIBLES---------------------------------------------------------------
   
   FlexProd.ColWidth(0) = 200
   FlexProd.ColWidth(1) = 1000
   FlexProd.ColWidth(2) = 16460
   FlexProd.ColWidth(3) = 1200
   FlexProd.ColWidth(4) = 1200
   FlexProd.ColWidth(5) = 2000
   FlexProd.ColWidth(6) = 3000
   
   FlexProd.TextMatrix(0, 1) = "Código"
   FlexProd.TextMatrix(0, 2) = "Descripcion"
   FlexProd.TextMatrix(0, 3) = "Stock"
   FlexProd.TextMatrix(0, 4) = "Stock Mínimo"
   FlexProd.TextMatrix(0, 5) = "Unidad de Medida"
   FlexProd.TextMatrix(0, 6) = "Sector"
   
   
     
   Set mRec = mObj.oEjecutarSelect("SELECT P.Codigo, P.Descripcion, P.Stock, P.Stock_Min, U.Descripcion AS UnidadMedida, S.Descripcion AS Sector " & _
     " From " & _
     " Producto P  " & _
     " Inner Join  " & _
     " UnidadMedida U ON P.CodUnidadMedida = U.Codigo  " & _
     " Inner Join  " & _
     " Sector S ON P.CodSector = S.Codigo " & _
     " where P.Fecha_Baja is null " & _
     " ORDER BY Codigo; ")
         
   If Not mRec.EOF Then
      mI = 1
      Do While Not mRec.EOF
         mI = mI + 1
         
         FlexProd.AddItem ""
         FlexProd.TextMatrix(mI, 1) = mRec!Codigo
         FlexProd.TextMatrix(mI, 2) = mRec!descripcion
         FlexProd.TextMatrix(mI, 3) = mRec!Stock
         FlexProd.TextMatrix(mI, 4) = mRec!Stock_Min
         FlexProd.TextMatrix(mI, 5) = mRec!UnidadMedida
         FlexProd.TextMatrix(mI, 6) = mRec!Sector
       
         mRec.MoveNext
      Loop
      FlexProd.RemoveItem 1
   End If
   mRec.Close
   
   '-FIN: GRILLA PRODUCTOS DISPONIBLES-----------------------------------------------------------------------------------------------------------------
   
   '--CARGO COMBO MOTIVOS EGRESO
   
   Set mRec = mObj.oTabla("MotivosEgreso", "")
   Do While Not mRec.EOF
      Combo1.AddItem "" & mRec!Codigo & " " & mRec!descripcion & ""
      mRec.MoveNext
   Loop
   mRec.Close
   
   '--FIN: CARGO COMBO MOTIVOS
   
  '-------------------------------------------------GRILLA PRODUCTOS CONSUMIDOS----------------------------------------------------------------------
   FlexEgreso.ColWidth(0) = 200
   FlexEgreso.ColWidth(1) = 1000
   FlexEgreso.ColWidth(2) = 12400
   FlexEgreso.ColWidth(3) = 1200
   FlexEgreso.ColWidth(4) = 3240
   FlexEgreso.ColWidth(5) = 7000
   FlexEgreso.ColWidth(6) = 0
   
   FlexEgreso.TextMatrix(0, 1) = "Código"
   FlexEgreso.TextMatrix(0, 2) = "Descripcion"
   FlexEgreso.TextMatrix(0, 3) = "Cantidad"
   FlexEgreso.TextMatrix(0, 4) = "Motivo"
   FlexEgreso.TextMatrix(0, 5) = "Observaciones"
   FlexEgreso.TextMatrix(0, 6) = "CodMotivo"
   
   
   FlexEgreso.ColAlignment(4) = 2
  '--------------------------------------------------FIN GRILLA PRODUCTOS CONSUMIDOS---------------------------------------------------------------
 
End Sub

Private Sub Command1_Click()
   Dim i As Integer

   If fValidaEgreso() Then
      ReDim Preserve vMovimientos(0 To contDim) As Movimiento
      
      vMovimientos(contDim).CodProducto = FlexProd.TextMatrix(mRenglon, 1)
      vMovimientos(contDim).Cantidad = CDbl(Replace(Trim(Text1.Text), ".", ","))
      
      FlexEgreso.AddItem vbTab & FlexProd.TextMatrix(mRenglon, 1) & vbTab & FlexProd.TextMatrix(mRenglon, 2) & vbTab & Text1.Text & vbTab & Combo1.Text & vbTab & Text2.Text & vbTab & Left(Combo1.Text, 2)
      If FlexEgreso.TextMatrix(1, 1) = "" Then
         FlexEgreso.RemoveItem 1
      End If
      contDim = contDim + 1
      Text1.Text = ""
   End If
   
End Sub

Private Sub Command2_Click(Index As Integer)
Dim i As Integer
Dim mI As Integer
Dim cantFilas As Integer

If Index = 0 Then
   If contDim > 0 Then
      cantFilas = FlexEgreso.Rows - 1
      
      For i = 1 To cantFilas
         mObj.xInsMovimiento FlexEgreso.TextMatrix(i, 1), "E", FlexEgreso.TextMatrix(i, 3), FlexEgreso.TextMatrix(i, 6), Trim(Right(MDI.mUser, 15)), FlexEgreso.TextMatrix(i, 5)
      Next
   
      For i = cantFilas To 1 Step -1
         If i <> 1 Then
            FlexEgreso.RemoveItem (i)
         Else
               FlexEgreso.Clear
               FlexEgreso.TextMatrix(0, 1) = "Código"
               FlexEgreso.TextMatrix(0, 2) = "Descripcion"
               FlexEgreso.TextMatrix(0, 3) = "Cantidad"
               FlexEgreso.TextMatrix(0, 4) = "Motivo"
               FlexEgreso.TextMatrix(0, 5) = "Observaciones"
               FlexEgreso.TextMatrix(0, 6) = "CodMotivo"
         End If
      Next
      
      Erase vMovimientos
      contDim = 0
      'TODO: Actualizar grilla productos (superior), en funcion del filtro
      
      cantFilas = FlexProd.Rows - 1
      
      For i = cantFilas To 1 Step -1
         If i <> 1 Then
            FlexProd.RemoveItem (i)
         Else
            FlexProd.Clear
            FlexProd.TextMatrix(0, 1) = "Código"
            FlexProd.TextMatrix(0, 2) = "Descripcion"
            FlexProd.TextMatrix(0, 3) = "Stock"
            FlexProd.TextMatrix(0, 4) = "Stock Mínimo"
            FlexProd.TextMatrix(0, 5) = "Unidad de Medida"
            FlexProd.TextMatrix(0, 6) = "Sector"
         End If
      Next
      
      
      mRenglon = 0
      
     Set mRec = mObj.oEjecutarSelect("SELECT P.Codigo, P.Descripcion, P.Stock, P.Stock_Min, U.Descripcion AS UnidadMedida, S.Descripcion AS Sector " & _
       " From " & _
       " Producto P  " & _
       " Inner Join  " & _
       " UnidadMedida U ON P.CodUnidadMedida = U.Codigo  " & _
       " Inner Join  " & _
       " Sector S ON P.CodSector = S.Codigo " & _
       " where P.Fecha_Baja is null " & _
       " ORDER BY Codigo; ")
     
      
      If Not mRec.EOF Then
      mI = 1
      Do While Not mRec.EOF
         mI = mI + 1
         
         FlexProd.AddItem ""
         FlexProd.TextMatrix(mI, 1) = mRec!Codigo
         FlexProd.TextMatrix(mI, 2) = mRec!descripcion
         FlexProd.TextMatrix(mI, 3) = mRec!Stock
         FlexProd.TextMatrix(mI, 4) = mRec!Stock_Min
         FlexProd.TextMatrix(mI, 5) = mRec!UnidadMedida
         FlexProd.TextMatrix(mI, 6) = mRec!Sector
       
         mRec.MoveNext
      Loop
      FlexProd.RemoveItem 1
   End If
   mRec.Close
   Text1.Text = ""
   Text2.Text = ""
   
   MsgBox "El ingreso se ha realizado exitosamente !!!", vbInformation, "Ingreso de Productos"
         
   Else
      MsgBox "Debe agregar al menos un producto en la grilla inferior", vbInformation, "Atención"
   End If
Else
   Unload Me
End If
   
End Sub

Private Sub FlexProd_Click()

Dim mI As Integer

'If FlexProd.MouseCol = 0 And FlexProd.MouseRow > 0 Then

If FlexProd.MouseRow > 0 Then

   If mRenglon <> 0 Then
      FlexProd.Row = mRenglon
      For mI = 0 To FlexProd.Cols - 1
         FlexProd.Col = mI
         FlexProd.CellBackColor = vbWhite
      Next
   End If
   
   mRenglon = FlexProd.MouseRow

   FlexProd.Row = mRenglon
   For mI = 0 To FlexProd.Cols - 1
      FlexProd.Col = mI
      FlexProd.CellBackColor = &H8000000D
   Next
Else
   mRenglon = 0
End If


End Sub

Private Sub Form_Load()
   Inicio
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
      If KeyAscii <> 46 Then
         KeyAscii = fNumeroKeyPress(KeyAscii)
      End If
End Sub

Private Function fValidaEgreso() As Boolean
   Dim mRet As Boolean
   Dim mMensajeError As String
   Dim j As Integer
   Dim mCantidadMovida As Double
   Dim mCantidaStock As Double
   Dim mCodProducto As String
   
   mRet = True
      
   If mRenglon = 0 Then
      mRet = False
      mMensajeError = "Debe seleccionar un producto de la grilla superior"
   End If
      
   If mRet Then
      If mRenglon <> 0 And FlexProd.TextMatrix(mRenglon, 1) = "" Then
         mRet = False
         mMensajeError = "Debe seleccionar un producto de la grilla superior"
      End If
   End If
      
      
   If mRet Then
      If Trim(Text1.Text) = "" Or Trim(Combo1.Text) = "" Or Trim(Text2.Text) = "" Then
         mRet = False
         mMensajeError = "Debe completar todos los datos"
      End If
   End If
      
   If mRet Then
      'If Not IsNumeric(Text1.Text) Then
      If Not IsNumeric(Replace(Text1.Text, ".", ",")) Then
         mRet = False
         mMensajeError = "La Cantidad ingresada no es un valor numérico"
      End If
   End If
      
   'Vvalido si el saldo del stock es insuficiente
   If mRet Then
   
      mCantidadMovida = 0
      mCodProducto = FlexProd.TextMatrix(mRenglon, 1)
   
      'mCantidaStock = (Replace(mObj.sTablaDescr("Producto", "Codigo = '" & mCodProducto & "'", 4), ",", "."))
      mCantidaStock = mObj.sTablaDescr("Producto", "Codigo = '" & mCodProducto & "'", 4)
      
      If contDim <> 0 Then
         For j = 0 To contDim - 1
           If vMovimientos(j).CodProducto = mCodProducto Then
              mCantidadMovida = mCantidadMovida + vMovimientos(j).Cantidad
           End If
         Next
      End If
      
      mCantidadMovida = mCantidadMovida + CDbl(Replace(Trim(Text1.Text), ".", ","))
      
      If mCantidaStock < mCantidadMovida Then
         mRet = False
         mMensajeError = "Stock disponible insuficiente"
      End If
   End If

   
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   fValidaEgreso = mRet
End Function


Private Sub Form_Unload(Cancel As Integer)
Set mObj = Nothing
Set mRec = Nothing
ShowMenu 12, True, False
End Sub

