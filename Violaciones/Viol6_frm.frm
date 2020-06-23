VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Viol6_frm 
   BackColor       =   &H00CECECE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulo de Búsqueda de Datos"
   ClientHeight    =   6390
   ClientLeft      =   4320
   ClientTop       =   330
   ClientWidth     =   11925
   ControlBox      =   0   'False
   Icon            =   "Viol6_frm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11925
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1695
      Left            =   1800
      TabIndex        =   25
      Top             =   1800
      Visible         =   0   'False
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2990
      _Version        =   327680
      Cols            =   4
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Editar Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   22
      Top             =   3840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   10
      Top             =   5880
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00CECECE&
      Caption         =   "Solo con Cartas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   9
      Top             =   5760
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlex2 
      Height          =   2775
      Left            =   4200
      TabIndex        =   12
      Top             =   3480
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4895
      _Version        =   327680
      Cols            =   10
      FixedCols       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlex1 
      Height          =   2775
      Left            =   4200
      TabIndex        =   11
      Top             =   360
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4895
      _Version        =   327680
      Cols            =   6
      FixedCols       =   0
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   8
      Top             =   4560
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   7
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Dbl-Click para ver historial"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   2400
      TabIndex        =   26
      Top             =   1260
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000FF&
      Caption         =   "Patente en Stand By desde el 15/04/2010"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   840
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00CECECE&
      Caption         =   "En una búsqueda, usar tecla % como comodín. Ej. (%TT%89)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   840
      TabIndex        =   23
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00CECECE&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   1
      Left            =   10200
      TabIndex        =   21
      Top             =   3195
      Width           =   675
   End
   Begin VB.Label Label3 
      BackColor       =   &H00CECECE&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   0
      Left            =   10200
      TabIndex        =   20
      Top             =   75
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "Detalle de Pasadas en Violación                                        Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   4275
      TabIndex        =   19
      Top             =   3240
      Width           =   5670
   End
   Begin VB.Label Label2 
      BackColor       =   &H00CECECE&
      Caption         =   "Detalle de Envíos de Cartas Documentos                            Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   4275
      TabIndex        =   18
      Top             =   120
      Width           =   5700
   End
   Begin VB.Label Label1 
      BackColor       =   &H00CECECE&
      Caption         =   "Provincia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   17
      Top             =   4320
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "C. Postal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   16
      Top             =   3600
      Width           =   780
   End
   Begin VB.Label Label1 
      BackColor       =   &H00CECECE&
      Caption         =   "Localidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   15
      Top             =   2880
      Width           =   840
   End
   Begin VB.Label Label1 
      BackColor       =   &H00CECECE&
      Caption         =   "Dirección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   14
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label Label1 
      BackColor       =   &H00CECECE&
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   1440
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "Patente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "Viol6_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObj As New clViolaciones
Dim mObjPea As New clPeaje
Dim mRec As New ADODB.Recordset
Dim mEstaciones(18) As String
Dim mI As Integer

Private Sub Form_Load()
Me.Height = 6795
Me.Width = 12050
sAlinearForm Me
Set mRec = mObjPea.oEstaciones("")
While Not mRec.EOF
   mEstaciones(mRec!CODIGO_ESTACION) = Trim(mRec!Descripcion_Estacion)
   mRec.MoveNext
Wend
mRec.Close
sTituloFlex
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mObj = Nothing
Set mObjPea = Nothing
Set mRec = Nothing
ShowMenu 5, True, False
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
   Case 0
      Me.MousePointer = 11
      If Text1(0).Visible And (Trim(Text1(0).Text) <> "" Or Trim(Text1(1).Text) <> "" Or Trim(Text1(2).Text) <> "") Then
         If Check1.Value = 1 Then
            If Trim(Text1(1).Text) = "" And Trim(Text1(2).Text) = "" Then
               Set mRec = mObj.oDistPatenteTabla("envios", Trim(Text1(0).Text), Trim(Text1(2).Text), Trim(Text1(1).Text))
            Else
               Set mRec = mObj.oDistPatenteDirEnv(Trim(Text1(0).Text), Trim(Text1(2).Text), Trim(Text1(1).Text))
            End If
         Else
            If Trim(Text1(1).Text) = "" And Trim(Text1(2).Text) = "" Then
               Set mRec = mObj.oDistPatenteTabla("Registros", Trim(Text1(0).Text), Trim(Text1(2).Text), Trim(Text1(1).Text))
            Else
               Set mRec = mObj.oDistPatenteTabla("direcciones", Trim(Text1(0).Text), Trim(Text1(2).Text), Trim(Text1(1).Text))
            End If
         End If
         If Not mRec.EOF Then
            Text1(0).Visible = False
            Text1(1).Enabled = False
            Text1(2).Enabled = False
            Text1(1).Text = ""
            Text1(2).Text = ""
            Combo1.Visible = True
            Do While Not mRec.EOF
               Combo1.AddItem mRec!patente
               mRec.MoveNext
            Loop
         Else
            MsgBox "No Existen Datos para esta Búsqueda.", vbInformation, sMessage
         End If
         mRec.Close
      Else
         Combo1.Clear
         Combo1.Visible = False
         For mI = 0 To Text1.UBound
            Text1(mI).Text = ""
         Next
         Text1(0).Visible = True
         Text1(1).Enabled = True
         Text1(2).Enabled = True
         Command2.Visible = False
         Label5.Visible = False
         Label6.Visible = False
      End If
      Me.MousePointer = 0
   Case 1
      Unload Viol6_frm
End Select
End Sub

Private Sub Command2_Click()
Viol8_frm.pPatente = Trim(Combo1.Text)
Viol8_frm.Show
Viol6_frm.Enabled = False
End Sub

Private Sub Combo1_Click()
Dim mColor As String
Dim mModelo As String
Dim mMarca As String
Me.MousePointer = 11
sMsgEspere Me, "Buscando datos...", True
For mI = 0 To Text1.UBound
   Text1(mI).Text = ""
Next
Set mRec = mObj.oDatosPatente(Trim(Combo1.Text))
If Not mRec.EOF Then
   Do While Not mRec.EOF
      Text1(1).Text = mRec!nombre
      Text1(2).Text = mRec!domicilio
      Text1(3).Text = mObj.sCampoDescrip("postal", "codigo='" & mRec!codpostal & "' and codpcia='" & mRec!codpcia & "'", 2)
      Text1(4).Text = mRec!codpostal
      Text1(5).Text = mObj.sCampoDescrip("provincias", "codigo='" & mRec!codpcia & "'", 1)
      mRec.MoveNext
   Loop
Else
   Text1(1).Text = "NO EXISTEN DATOS..."
End If
mRec.Close
Command2.Visible = True
Set mRec = mObj.oEjecutarSelect("SELECT * FROM regpagos WHERE patente = '" & Combo1.Text & "'")
If Not mRec.EOF Then
   Do While Not mRec.EOF
      If Right(Trim(mRec!estado), 1) = "1" Then
         Label5.Caption = "Stand By - " & mRec!Fecha
         Label5.BackColor = "&H000000FF"
      Else
         Label5.Caption = "Activo - " & mRec!Fecha
         Label5.BackColor = "&H0000FF00"
      End If
      mRec.MoveNext
   Loop
   Label5.Visible = True
   Label6.Visible = True
End If
mRec.Close
'VIOLACIONES POR ESTACION
sBorraFlexDatos Viol6_frm.MSFlex2
Set mRec = mObj.oRegistrosPatente(Trim(Combo1.Text))
Do While Not mRec.EOF
   mMarca = mObj.sCampoDescrip("marcas", "codigo='" & mRec!CodMarca & "'", 1)
   mModelo = mObj.sCampoDescrip("modelos", "codigo='" & mRec!modelo & "' and codmarca='" & mRec!CodMarca & "'", 2)
   mColor = NVL(mRec!Color, "")
   If Len(mColor) <= 2 Then
      mColor = mObj.sCampoDescrip("colores", "codigo='" & mColor & "'", 1)
   End If
   MSFlex2.AddItem "" & vbTab & " " & mRec!Fecha & vbTab & " " & mRec!Hora & vbTab & " " & mRec!Estacion & "-" & mEstaciones(Int(Val(mRec!Estacion))) & vbTab & " " & mRec!Via & vbTab & " " & mMarca & vbTab & " " & mModelo & vbTab & mColor & vbTab & mRec!Tipo & vbTab & mRec!pago
   mRec.MoveNext
Loop
mRec.Close
If MSFlex2.Rows > 2 Then
   MSFlex2.RemoveItem 1
End If
If MSFlex2.TextMatrix(1, 1) <> "" Then
   sSetFlex2Colors Viol6_frm.MSFlex2, &HEFEFF1, &HF4E7E1
   sSetFlexNroFila Viol6_frm.MSFlex2, 0
   Label3(1).Caption = MSFlex2.Rows - 1
Else
   Label3(1).Caption = "0"
End If
'ENVIOS DE CARTAS DOCUMENTOS
Set mRec = mObj.oCDporPatente(Trim(Combo1.Text))
Do While Not mRec.EOF
   MSFlex1.AddItem "" & vbTab & mRec!Fecha & vbTab & " " & mRec!NROCARTA & vbTab & mRec!descripcion & vbTab & " " & mRec!OBS & vbTab & mRec!Tipo
   mRec.MoveNext
Loop
mRec.Close
If MSFlex1.Rows > 2 Then
   MSFlex1.RemoveItem 1
End If
If MSFlex1.TextMatrix(1, 1) <> "" Then
   sSetFlex2Colors Viol6_frm.MSFlex1, &HEFEFF1, &HF4E7E1
   sSetFlexNroFila Viol6_frm.MSFlex1, 0
   Label3(0).Caption = MSFlex1.Rows - 1
Else
   Label3(0).Caption = "0"
End If
Me.MousePointer = 0
sMsgEspere Me, "", False
End Sub

Private Sub sTituloFlex()
With MSFlex1
   .ColWidth(0) = 500
   .ColWidth(1) = 1000
   .ColWidth(2) = 2100
   .ColWidth(3) = 2000
   .ColWidth(4) = 3800
   .ColWidth(5) = 400
   .TextMatrix(0, 0) = "N°"
   .TextMatrix(0, 1) = "Fecha"
   .TextMatrix(0, 2) = "N° Carta"
   .TextMatrix(0, 3) = "Entrega"
   .TextMatrix(0, 4) = "Observaciones"
   .TextMatrix(0, 5) = "Tipo"
End With
With MSFlex2
   .ColWidth(0) = 500
   .ColWidth(1) = 1000
   .ColWidth(2) = 700
   .ColWidth(3) = 1500
   .ColWidth(4) = 500
   .ColWidth(5) = 1000
   .ColWidth(6) = 1200
   .ColWidth(7) = 900
   .ColWidth(8) = 400
   .ColWidth(9) = 500
   .TextMatrix(0, 0) = "N°"
   .TextMatrix(0, 1) = "Fecha"
   .TextMatrix(0, 2) = "Hora"
   .TextMatrix(0, 3) = "Estación"
   .TextMatrix(0, 4) = "Vía"
   .TextMatrix(0, 5) = "Marca"
   .TextMatrix(0, 6) = "Modelo"
   .TextMatrix(0, 7) = "Color"
   .TextMatrix(0, 8) = "Tipo"
   .TextMatrix(0, 9) = "Pago"
   For mI = 0 To 7
      .Col = mI
      .CellFontBold = True
   Next
End With
With MSFlexGrid1
   .ColWidth(0) = 200
   .ColWidth(1) = 1900
   .ColWidth(2) = 800
   .ColWidth(3) = 10000
   .TextMatrix(0, 0) = ""
   .TextMatrix(0, 1) = "Fecha"
   .TextMatrix(0, 2) = "Estado"
   .TextMatrix(0, 3) = "Observaciones"
End With
End Sub

Private Sub Label5_DblClick()
sBorraFlexDatos MSFlexGrid1
Set mRec = mObj.oEjecutarSelect("SELECT * FROM regpagos WHERE patente = '" & Combo1.Text & "'")
If Not mRec.EOF Then
   Do While Not mRec.EOF
      MSFlexGrid1.AddItem "" & vbTab & " " & mRec!Fecha & vbTab & " " & IIf(Right(Trim(mRec!estado), 1) = "1", "Stand By", "Activo") & vbTab & " " & mRec!OBS
      mRec.MoveNext
   Loop
End If
sSetFlex2Colors MSFlexGrid1, &HEFEFF1, &HF4E7E1
MSFlexGrid1.RemoveItem 1
MSFlexGrid1.Visible = True
End Sub

Private Sub MSFlex2_DblClick()
If MSFlex2.Col = 9 And MSFlex2.Row >= 1 Then
   mObj.UpdPagos IIf(MSFlex2.TextMatrix(MSFlex2.Row, 9) = "S", "", "S"), MSFlex2.TextMatrix(MSFlex2.Row, 1), Trim(MSFlex2.TextMatrix(MSFlex2.Row, 2)), Combo1.Text
   MSFlex2.TextMatrix(MSFlex2.Row, 9) = IIf(MSFlex2.TextMatrix(MSFlex2.Row, 9) = "S", "", "S")
   mObj.InsLogPagos Trim(Right(MDI.mUser, 15)), Now, Combo1.Text, MSFlex2.TextMatrix(MSFlex2.Row, 9), MSFlex2.TextMatrix(MSFlex2.Row, 1), Trim(MSFlex2.TextMatrix(MSFlex2.Row, 2)), Trim(Left(MSFlex2.TextMatrix(MSFlex2.Row, 3), 3))
End If
End Sub

Private Sub MSFlexGrid1_DblClick()
MSFlexGrid1.Visible = False
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
   Case 0
      If KeyAscii >= 97 And KeyAscii <= 122 Then
         KeyAscii = KeyAscii - 32
      End If
      If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 44 And KeyAscii <= 57)) And KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 209 And KeyAscii <> 241 And KeyAscii <> 37 Then
         KeyAscii = 0
      Else
         If KeyAscii = 241 Then
            KeyAscii = 209
         End If
      End If
End Select
End Sub

