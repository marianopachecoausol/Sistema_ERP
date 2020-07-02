VERSION 5.00
Begin VB.Form MantElect24 
   Caption         =   "Reemprimir Ord. Trabajo Abierta"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4455
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   2790
   ScaleWidth      =   4455
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "OT - FECHA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "MantElect24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mObj As New clMantElect
Dim mRec As New ADODB.Recordset
Dim XLS As EXCEL.Application
Dim mi As Integer
Dim mFechaEjec As Date

Dim mLinea As Integer

Private Sub Command1_Click(Index As Integer)
  ' Dim mCodActivo As String
  ' Dim mDescActivo As String
   
   Dim mIdOT As Integer
  
   If Index = 0 Then
      If Combo1.Text <> "" Then
        
         mIdOT = Left(Combo1.Text, 10)
        
         sMsgEspere Me, "Procesando datos...", True
         
         Set XLS = CreateObject("Excel.Application")
         sPlanilla1 mIdOT

         sMsgEspere Me, "", False
         XLS.Application.Visible = True
      Else
         MsgBox "Debe seleccionar una Orden de Trabajo !!!!", vbCritical, sMessage
      End If
   Else
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Me.Width = 4695
   Me.Height = 3360
   sAlinearForm Me
   
    If mObj.esSupervisorElectrico(Trim(Right(MDI.mUser, 20))) Then
      Set mRec = mObj.oEjecutarSelect("SELECT DISTINCT CONVERT( CONCAT(LPAD(O.IdOT,10,'0'),' - ',Date_Format(O.Fecha,'%d/%m/%Y')),char(23)) AS OT_Fecha, O.Fecha " & _
         " FROM MantElect.OT_H O " & _
         " Inner Join " & _
         " OT_Partes OP ON O.IdOT = OP.IdOT " & _
         " Inner Join " & _
         " Registros R ON OP.Parte = R.Parte " & _
         " Where SectorAire = 0 " & _
         " AND O.FechaFin IS NULL " & _
         " ORDER BY O.IdOT DESC; ")
   Else
      Set mRec = mObj.oEjecutarSelect("SELECT DISTINCT CONVERT( CONCAT(LPAD(O.IdOT,10,'0'),' - ',Date_Format(O.Fecha,'%d/%m/%Y')),char(23)) AS OT_Fecha, O.Fecha " & _
         " FROM MantElect.OT_H O " & _
         " Inner Join " & _
         " OT_Partes OP ON O.IdOT = OP.IdOT " & _
         " Inner Join " & _
         " Registros R ON OP.Parte = R.Parte " & _
         " Where SectorAire = 1 " & _
         " AND O.FechaFin IS NULL " & _
         " ORDER BY O.IdOT DESC; ")
   End If
   
   Do While Not mRec.EOF
      Combo1.AddItem mRec!OT_Fecha
      mRec.MoveNext
   Loop
   mRec.Close
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mObj = Nothing
   Set mRec = Nothing
   Set XLS = Nothing
   ShowMenu 12, True, False
End Sub

Private Sub sPlanilla1(NroOT As Integer)
    Dim FechaOT As Date
    Dim Supervisor As String
    
    
   'Dim mi As Integer
   Dim mLineasXpagina As Integer
   Dim primerColumna As Boolean
   Dim mj As Integer
   mLinea = 1
   mLineasXpagina = 81
   'mi = 250
   
   
    Set mRec = mObj.oEjecutarSelect("SELECT DISTINCT Fecha, CONCAT( U.apellido, ', ',U.nombres) as Supervisor " & _
                                          "FROM OT_H OH " & _
                                              "Inner Join " & _
                                          "loguser.usuarios U ON OH.CodUsuario = U.CodUsuario " & _
                                      "WHERE IdOT = '" & NroOT & "';")
                                
    If Not mRec.EOF Then
        FechaOT = mRec!Fecha
        Supervisor = mRec!Supervisor
    
    End If
    mRec.Close
     
   With XLS
      .WorkBooks.Add
      .Worksheets(1).Select
      .ActiveWindow.DisplayGridlines = False
      .Worksheets(1).Name = "Orden de Trabajo"
     
      .Columns("A:A").ColumnWidth = 1.14 '
      .Columns("B:B").ColumnWidth = 6.86 '
      .Columns("C:C").ColumnWidth = 24.29 '
      .Columns("D:D").ColumnWidth = 9.71 '
      .Columns("F:F").ColumnWidth = 10.29 '
      .Columns("G:G").ColumnWidth = 10.29 '
      .Columns("I:I").ColumnWidth = 9.86 '
      .Columns("J:J").ColumnWidth = 1.14 '

      .Range("B1:J500").Select
      .Selection.Font.Size = 7
      .Selection.Font.Bold = True
      .Selection.RowHeight = 10.5

'---------------------------------ENCABEZADO HOJA-------------------------------------------------------
      .Cells(mLinea, 2).Formula = "AUTOPISTAS DEL SOL S.A."
      .Cells(mLinea + 1, 4).Formula = "PLANILLA DE ORDEN DE TRABAJO"
      
      .Cells(mLinea + 3, 2).Formula = "Fecha: " & FechaOT
      .Cells(mLinea + 4, 2).Formula = "Tipo Tarea"
      .Cells(mLinea + 5, 2).Formula = "Supervisor: " & Supervisor
      
      .Cells(mLinea + 3, 8).Formula = "Nº OT"
      .Cells(mLinea + 4, 8).Formula = "Hora Inicio"
      .Cells(mLinea + 5, 8).Formula = "Hora Fin"
      .Cells(mLinea + 3, 9).Formula = NroOT

      .Range("H4:H6").Select
      .Selection.Interior.ColorIndex = 15

      .Range("H" & (mLinea + 3) & ":I" & (mLinea + 5)).Select
      With .Selection.Borders(xlBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      mLinea = mLinea + 8
'---------------------------------ENCABEZADO TECNICOS---------------------------------------------------
       If mLinea Mod mLineasXpagina = 0 Then
         MsgBox "FIN: MOD = 0 Encabezado Tecnicos"
         'Repetir lo del else
       Else
          .Cells(mLinea, 4).Formula = "TECNICOS QUE INTERVIENEN"
         
         .Range("B" & mLinea & ":H" & mLinea).Select
         .Selection.Interior.ColorIndex = 15
         With .Selection.Borders(xlBottom)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlTop)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlEdgeRight)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
      End If
'-----------------------------------------------------------------------------------------------------
'---------------------------------DETALLE TECNICOS----------------------------------------------------
      primerColumna = 1
      Set mRec = mObj.oEjecutarSelect("SELECT Codigo, Descripcion " & _
                                          "FROM OT_MO_Tecnicos O " & _
                                              "Inner Join " & _
                                          "MO_Tecnicos M ON O.CodMO_Tecnico = M.Codigo " & _
                                      "WHERE IdOT = '" & NroOT & "';")
                                
      mLinea = mLinea + 1
      Do While Not mRec.EOF
         
         'if linea mod then
            'Imprimir encaabezado
            'Incrementar mHoja
            'Poscionar mLinea donde corresponda
         'End If
         
         
         .Range("B" & mLinea & ":H" & mLinea).Select
         With .Selection.Borders(xlBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlEdgeRight)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         .Range("E" & mLinea & ":E" & mLinea).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         If primerColumna Then
            .Cells(mLinea, 2).Formula = NVL(mRec!descripcion, "")
            primerColumna = False
         Else
            .Cells(mLinea, 5).Formula = NVL(mRec!descripcion, "")
            primerColumna = True
            mLinea = mLinea + 1
         End If
   
         mRec.MoveNext
      Loop
      mRec.Close
'-----------------------------------------------------------------------------------------------------



'---------------------------------ENCABEZADO VEHICULOS------------------------------------------------

      mLinea = mLinea + 2
         'if (mlinea mod = 0) or (mLinea+1 mod = 0) then
            'Imprimir encaabezado (deberia modificar el valor mLinea)
            'Incrementar mHoja
            'Poscionar mLinea donde corresponda
         'End If
      .Range("B" & mLinea & ":H" & (mLinea + 1)).Select
      .Selection.Interior.ColorIndex = 15

      .Range("B" & mLinea & ":H" & (mLinea + 1)).Select
      With .Selection.Borders(xlBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      .Range("E" & (mLinea + 1) & ":H" & (mLinea + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      .Range("G" & (mLinea + 1) & ":G" & (mLinea + 1)).Select
       With .Selection.Borders(xlEdgeLeft)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
       End With
      
      .Cells(mLinea, 4).Formula = "VEHICULOS QUE INTERVIENEN"
      mLinea = mLinea + 1
      .Cells(mLinea, 2).Formula = "Vehículo"
      .Cells(mLinea, 5).Formula = "Km Inicial"
      .Cells(mLinea, 7).Formula = "Km Final"

'-----------------------------------------------------------------------------------------------------


'---------------------------------DETALLE VEHICULOS------------------------------------------------
      mLinea = mLinea + 1
      Set mRec = mObj.oEjecutarSelect("SELECT Codigo,Descripcion FROM " & _
                                          "OT_Vehiculos O " & _
                                              "Inner Join " & _
                                          "Vehiculos V ON O.CodVehiculo = Codigo " & _
                                      "WHERE IdOT = '" & NroOT & "'; ")
                                
      Do While Not mRec.EOF
      
         'if (mlinea mod = 0)  then
            'Imprimir encaabezado (deberia modificar el valor mLinea)
            'Incrementar mHoja
            'Poscionar mLinea donde corresponda
         'End If
         
         
         .Range("B" & mLinea & ":H" & mLinea).Select
         With .Selection.Borders(xlBottom)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlTop)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlEdgeRight)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
      
         .Range("E" & mLinea & ":E" & mLinea).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
      
         .Range("G" & mLinea & ":G" & mLinea).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
         
         With XLS
            .Cells(mLinea, 2).Formula = NVL(mRec!descripcion, "")
         End With
         mRec.MoveNext
         mLinea = mLinea + 1
      Loop
      mRec.Close
'-----------------------------------------------------------------------------------------------------



'---------------------------------ENCABEZADO TAREAS--------------------------------------------------
      
      
      
      'if (mlinea mod = 0) or (mLinea+1 mod = 0) or (mLinea+2 mod = 0) then
            'Imprimir encaabezado (deberia modificar el valor mLinea)
            'Incrementar mHoja
            'Poscionar mLinea donde corresponda ''mLinea = mLinea + 2
      'End If
      
      mLinea = mLinea + 2 'Borrar cuando descomente lo de arriba.
      
      .Cells(mLinea, 5).Formula = "TAREAS"
      .Cells(mLinea + 1, 2).Formula = "Parte"
      .Cells(mLinea + 1, 3).Formula = "Lugar"
      .Cells(mLinea + 1, 4).Formula = "Descripcion"
      .Cells(mLinea + 1, 9).Formula = "¿Finalizado?"

      .Range("B" & mLinea & ":I" & (mLinea + 1)).Select
      .Selection.Interior.ColorIndex = 15

      .Range("B" & mLinea & ":I" & (mLinea + 1)).Select
      With .Selection.Borders(xlBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      .Range("C" & (mLinea + 1) & ":C" & (mLinea + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      .Range("D" & (mLinea + 1) & ":D" & (mLinea + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      .Range("I" & (mLinea + 1) & ":I" & (mLinea + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
'-----------------------------------------------------------------------------------------------------
'---------------------------------DETALLE TAREAS------------------------------------------------------
      mLinea = mLinea + 2
      Set mRec = mObj.oEjecutarSelect("SELECT R.Parte,R.CodEdificio, R.Descripcion, Length(R.Descripcion) lenDesc " & _
                                          "FROM " & _
                                          "OT_Partes O " & _
                                              "Inner Join " & _
                                          "Registros R ON O.Parte = R.Parte " & _
                                          "WHERE IDOT = '" & NroOT & "' " & _
                                          "ORDER BY R.parte; ")
                                
      Do While Not mRec.EOF
         'if (mlinea mod = 0)  then
            'Imprimir encaabezado (deberia modificar el valor mLinea)
            'Incrementar mHoja
            'Poscionar mLinea donde corresponda
         'End If
         
         .Range("B" & mLinea & ":I" & mLinea).Select
         With .Selection.Borders(xlBottom)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlTop)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With

         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With

         With .Selection.Borders(xlEdgeRight)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         .Range("C" & mLinea & ":C" & mLinea).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
      
         .Range("D" & mLinea & ":D" & mLinea).Select
         If mRec!lenDesc > 75 Then
            .Selection.Font.Size = 6
         Else
         .Selection.Font.Size = 7
         End If
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With

         .Range("I" & mLinea & ":I" & mLinea).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With

         With XLS
            .Cells(mLinea, 2).Formula = NVL(mRec!Parte, "")
            .Cells(mLinea, 3).Formula = NVL(mRec!CodEdificio, "")
            .Cells(mLinea, 4).Formula = NVL(mRec!descripcion, "")
         End With
         mRec.MoveNext
         mLinea = mLinea + 1
      Loop
      mRec.Close
'-----------------------------------------------------------------------------------------------------
'---------------------------------ENCABEZADO SUBRUBROS------------------------------------------------
      'if (mlinea mod = 0) or (mLinea+1 mod = 0) or (mLinea+2 mod = 0) then
            'Imprimir encaabezado (deberia modificar el valor mLinea)
            'Incrementar mHoja
            'Poscionar mLinea donde corresponda ''mLinea = mLinea + 2
      'End If
      
      mLinea = mLinea + 2 'Borrar cuando descomente lo de arriba.
      
      
      .Cells(mLinea, 5).Formula = "FALLAS"
      .Cells(mLinea + 1, 2).Formula = "Subrubro"
      .Cells(mLinea + 1, 6).Formula = "Subrubro"

      .Range("B" & mLinea & ":I" & (mLinea + 1)).Select
      .Selection.Interior.ColorIndex = 15

      .Range("B" & mLinea & ":I" & (mLinea + 1)).Select
      With .Selection.Borders(xlBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      .Range("E" & mLinea + 1 & ":E" & (mLinea + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
      End With

      .Range("I" & mLinea + 1 & ":I" & (mLinea + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
'-----------------------------------------------------------------------------------------------------
'---------------------------------DETALLE SUBRUBROS-----------------------------------------------
      mLinea = mLinea + 2
      Set mRec = mObj.oEjecutarSelect("SELECT S.Codigo,S.Descripcion FROM " & _
                                       "SubRubros S " & _
                                          "Inner Join " & _
                                       "OT_Subrubros O ON O.CodSubrubro = S.Codigo " & _
                                       "WHERE IDOT = '" & NroOT & "' " & _
                                       "ORDER BY S.Descripcion; ")
                                
      primerColumna = True
      Do While Not mRec.EOF
      
         'if (mlinea mod = 0)  then
            'Imprimir encaabezado (deberia modificar el valor mLinea)
            'Incrementar mHoja
            'Poscionar mLinea donde corresponda
         'End If
   
         If primerColumna Then
            .Range("B" & mLinea & ":I" & mLinea).Select
            With .Selection.Borders(xlBottom)
              .LineStyle = xlContinuous
              .Weight = xlThin
              .ColorIndex = xlAutomatic
            End With
      
            With .Selection.Borders(xlTop)
              .LineStyle = xlContinuous
              .Weight = xlThin
              .ColorIndex = xlAutomatic
            End With
      
            With .Selection.Borders(xlEdgeLeft)
              .LineStyle = xlContinuous
              .Weight = xlThin
              .ColorIndex = xlAutomatic
            End With
      
            With .Selection.Borders(xlEdgeRight)
              .LineStyle = xlContinuous
              .Weight = xlThin
              .ColorIndex = xlAutomatic
            End With
         
            .Range("E" & mLinea & ":E" & mLinea).Select
            With .Selection.Borders(xlEdgeLeft)
              .LineStyle = xlContinuous
              .Weight = xlThin
              .ColorIndex = xlAutomatic
            End With
         
            .Range("F" & mLinea & ":F" & mLinea).Select
            With .Selection.Borders(xlEdgeLeft)
              .LineStyle = xlContinuous
              .Weight = xlMedium
              .ColorIndex = xlAutomatic
            End With
      
            .Range("I" & mLinea & ":I" & mLinea).Select
            With .Selection.Borders(xlEdgeLeft)
              .LineStyle = xlContinuous
              .Weight = xlThin
              .ColorIndex = xlAutomatic
            End With
            
            .Cells(mLinea, 2).Formula = NVL(mRec!descripcion, "")
            primerColumna = False
         Else
            .Cells(mLinea, 6).Formula = NVL(mRec!descripcion, "")
            primerColumna = True
            mLinea = mLinea + 1
         End If
         mRec.MoveNext
      Loop
      mRec.Close
'-----------------------------------------------------------------------------------------------------




'---------------------------------ENCABEZADO Materiales-----------------------------------------------
      
      'if (mlinea mod = 0) or (mLinea+1 mod = 0) or (mLinea+2 mod = 0) then
            'Imprimir encaabezado (deberia modificar el valor mLinea)
            'Incrementar mHoja
            'Poscionar mLinea donde corresponda ''mLinea = mLinea + 2
      'End If
      
      mLinea = mLinea + 2 'Borrar cuando descomente lo de arriba.
      
      
      .Cells(mLinea, 4).Formula = "                 MATERIALES"
      .Cells(mLinea + 1, 2).Formula = "Cód.Sap"
      .Cells(mLinea + 1, 3).Formula = "Descripción"
      '.Cells(mi + 1, 6).Formula = "Consumido"
      .Cells(mLinea + 1, 7).Formula = "Consumido"
      '.Cells(mi + 1, 7).Formula = "U.Medida"
      .Cells(mLinea + 1, 8).Formula = "U.Medida"
      
      .Range("B" & mLinea & ":H" & (mLinea + 1)).Select
      .Selection.Interior.ColorIndex = 15

      .Range("B" & mLinea & ":H" & (mLinea + 1)).Select
      With .Selection.Borders(xlBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      .Range("C" & (mLinea + 1) & ":C" & (mLinea + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      '.Range("F" & (mi + 1) & ":F" & (mi + 1)).Select
      .Range("G" & (mLinea + 1) & ":G" & (mLinea + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      '.Range("G" & (mi + 1) & ":G" & (mi + 1)).Select
      .Range("H" & (mLinea + 1) & ":H" & (mLinea + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
'-----------------------------------------------------------------------------------------------------


'---------------------------------DETALLE Materiales--------------------------------------------------
      mLinea = mLinea + 2
'      Set mRec = mObj.oEjecutarSelect("SELECT  idMov,  M.Fecha,  P.CodigoSap,  P.Descripcion,  Stock,  UM.Descripcion AS UnidadMedidad FROM " & _
'                                          "Inventario.Movimientos2 M " & _
'                                              "Inner Join " & _
'                                          "Inventario.Producto P ON M.CodProducto = P.Codigo " & _
'                                              "Inner Join " & _
'                                          "Inventario.UnidadMedida UM ON P.CodUnidadMedida = UM.Codigo " & _
'                                              "Inner Join " & _
'                                          "Inventario.Ubicaciones U ON  M.CodUbicacion = U.Codigo " & _
'                                             "Inner Join " & _
'                                          "Vehiculos V ON U.Codigo = V.CodUbicacion " & _
'                                             "Inner Join " & _
'                                          "OT_Vehiculos OV ON OV.CodVehiculo = V.Codigo " & _
'                                          "WHERE M.Fecha = (SELECT MAX(Fecha) " & _
'                                          "                From Inventario.Movimientos2 " & _
'                                          "                WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion) " & _
'                                          "and OV.IDOT = '" & NroOT & "' and stock > 0; ")
                                          
                                          
                                          
      Set mRec = mObj.oEjecutarSelect(" SELECT P.CodigoSap,P.descripcion,UM.Descripcion AS UnidadMedidad " & _
                                       " From " & _
                                       " Matriz_Reposicion_Ubicaciones MR " & _
                                       " Inner Join " & _
                                       " Inventario.Ubicaciones U ON U.Codigo = MR.CodUbicacion " & _
                                       " Inner Join " & _
                                       " Vehiculos V ON V.CodUbicacion = U.Codigo " & _
                                       " Inner Join " & _
                                       " OT_Vehiculos OTV ON OTV.CodVehiculo = V.Codigo " & _
                                       " Inner Join " & _
                                       " Inventario.Producto P ON P.Codigo = MR.CodProducto " & _
                                       " Left Join " & _
                                       " Inventario.Movimientos2 M ON M.CodProducto = MR.CodProducto AND M.CodUbicacion = MR.CodUbicacion " & _
                                       " Inner Join " & _
                                       " Inventario.UnidadMedida UM ON UM.Codigo = P.CodUnidadMedida " & _
                                       " Where OTV.IdOT = '" & NroOT & "' " & _
                                       " AND M.Fecha = (SELECT MAX(Fecha) " & _
                                       " From Inventario.Movimientos2 " & _
                                       " WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion) " & _
                                       " AND MR.FechaHasta = '0000-00-00 00:00:00'; ")
                                          
                                          
                                          
                                
  

      Do While Not mRec.EOF
      
         'if (mlinea mod = 0)  then
            'Imprimir encaabezado (deberia modificar el valor mLinea)
            'Incrementar mHoja
            'Poscionar mLinea donde corresponda
         'End If
      
         .Range("B" & mLinea & ":H" & mLinea).Select
         With .Selection.Borders(xlBottom)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlTop)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlEdgeRight)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
      
         .Range("C" & mLinea & ":C" & mLinea).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
      
         '.Range("F" & mi & ":F" & mi).Select
         .Range("G" & mLinea & ":G" & mLinea).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         '.Range("G" & mi & ":G" & mi).Select
         .Range("H" & mLinea & ":H" & mLinea).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
      
         With XLS
            .Cells(mLinea, 2).Formula = NVL(mRec!CodigoSap, "")
            .Cells(mLinea, 3).Formula = NVL(mRec!descripcion, "")
            .Cells(mLinea, 8).Formula = NVL(mRec!UnidadMedidad, "")
         End With
         mRec.MoveNext
         mLinea = mLinea + 1
      Loop
      mRec.Close
   
'-----------------------------------------------------------------------------------------------------
 
 
'----------------------------------------------OBSERVACIONES------------------------------------------
      mLinea = mLinea + 2
      
'      For mj = mLinea To mLinea + 10
'         If mLinea Mod 81 = 0 Then
'            mEsCorte = True
'            mj = 9999
'         End If
'      Next
'
'      If mEsCorte Then
'         'imprimirEncabezado
'      Else
'         'mLinea = mj
'      End If
'
      
      
      .Cells(mLinea, 2).Formula = "OBSERVACIONES"
      mLinea = mLinea + 1
      .Range("B" & mLinea & ":I" & (mLinea + 4)).Select
    '  .Selection.RowHeight = 16.5
      With .Selection.Borders(xlBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

'-----------------------------------------------------------------------------------------------------

 
'----------------------------------------------FIRMAS-------------------------------------------------
      mLinea = mLinea + 8
      .Cells(mLinea, 3).Formula = "              SUPERVISOR"
      .Cells(mLinea, 6).Formula = "     ENCARGADO BODEGA"
      
      .Range("C" & mLinea & ":C" & mLinea).Select
      With .Selection.Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      .Range("F" & mLinea & ":G" & mLinea).Select
      With .Selection.Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
End With

'-----------------------------------------------------------------------------------------------------
 
 
''  Configuracion de margenes.
'   With ActiveSheet.PageSetup
'      .LeftMargin = Application.CentimetersToPoints(0)
'      .RightMargin = Application.CentimetersToPoints(0)
'      .TopMargin = Application.CentimetersToPoints(0)
'      .BottomMargin = Application.CentimetersToPoints(0)
'   End With
'
   
   
   '  Configuracion de margenes.
'   ActiveSheet.PageSetup.LeftMargin = Application.CentimetersToPoints(0)
'   ActiveSheet.PageSetup.RightMargin = Application.CentimetersToPoints(0)
'   ActiveSheet.PageSetup.TopMargin = Application.CentimetersToPoints(0)
'   ActiveSheet.PageSetup.BottomMargin = Application.CentimetersToPoints(0)
   
End Sub






