VERSION 5.00
Begin VB.MDIForm MDI 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   Icon            =   "MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   MousePointer    =   1  'Arrow
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu RNov_Arch 
      Caption         =   "Archivos"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu RNov_ArcSub 
         Caption         =   "Emisoras"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu RNov_ArcSub 
         Caption         =   "Origen"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu RNov_ArcSub 
         Caption         =   "Otros"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu RNov_ArcSub 
         Caption         =   "Patrulleros"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu RNov_ArcSub 
         Caption         =   "Rutinas"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu RNov_ArcSub 
         Caption         =   "Servicios"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu RNov_ArcSub 
         Caption         =   "Tipo Móvil"
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu RNov_ArcSub 
         Caption         =   "Tipo Otros"
         Enabled         =   0   'False
         Index           =   7
      End
      Begin VB.Menu RNov_ArcSub 
         Caption         =   "Trabajos"
         Enabled         =   0   'False
         Index           =   8
      End
      Begin VB.Menu RNov_ArcSub 
         Caption         =   "Turnos"
         Enabled         =   0   'False
         Index           =   9
      End
      Begin VB.Menu RNov_ArcSub 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu RNov_ArcSub 
         Caption         =   "Blanquear Partes"
         Enabled         =   0   'False
         Index           =   11
      End
      Begin VB.Menu RNov_ArcSub 
         Caption         =   "Generar Códigos"
         Enabled         =   0   'False
         Index           =   12
      End
   End
   Begin VB.Menu RNov_Nove 
      Caption         =   "Novedad"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu RNov_NovRegi 
         Caption         =   "Registrar"
         Enabled         =   0   'False
      End
      Begin VB.Menu RNov_NovVer 
         Caption         =   "Ver"
         Enabled         =   0   'False
      End
      Begin VB.Menu RNov_NovModi 
         Caption         =   "Modificar"
         Enabled         =   0   'False
      End
      Begin VB.Menu RNov_NovDepu 
         Caption         =   "Depurar"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu RNov_Repo 
      Caption         =   "Reportes"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu RNov_RepSub 
         Caption         =   "Novedades de Móvil"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu RNov_RepSub 
         Caption         =   "Km x Móvil"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu RNov_RepSub 
         Caption         =   "Arribos"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu RNov_RepSub 
         Caption         =   "Códigos de Ambulancia"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu RNov_RepSub 
         Caption         =   "Emisoras"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu RNov_RepSub 
         Caption         =   "Origen"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu RNov_RepSub 
         Caption         =   "Novedades"
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu RNov_RepSub 
         Caption         =   "Retiro de Objetos"
         Enabled         =   0   'False
         Index           =   7
      End
      Begin VB.Menu RNov_RepSub 
         Caption         =   "Cuadro"
         Enabled         =   0   'False
         Index           =   8
      End
      Begin VB.Menu RNov_RepSub 
         Caption         =   "Rutinas"
         Enabled         =   0   'False
         Index           =   9
         Begin VB.Menu RNov_RepRutSub 
            Caption         =   "Resumen"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu RNov_RepRutSub 
            Caption         =   "Por Fecha"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu RNov_RepRutSub 
            Caption         =   "Detalle"
            Enabled         =   0   'False
            Index           =   2
         End
      End
      Begin VB.Menu RNov_RepSegu 
         Caption         =   "Seguimientos"
         Enabled         =   0   'False
         Begin VB.Menu RNov_RepSegSub 
            Caption         =   "Por Móvil"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu RNov_RepSegSub 
            Caption         =   "Por Móvil (xls)"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu RNov_RepSegSub 
            Caption         =   "Por Evento"
            Enabled         =   0   'False
            Index           =   2
         End
      End
      Begin VB.Menu RNov_RepServ 
         Caption         =   "Servicios"
         Enabled         =   0   'False
         Begin VB.Menu RNov_RepSerSub 
            Caption         =   "Total por Tipo/Móvil"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu RNov_RepSerSub 
            Caption         =   "Por Progresiva"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu RNov_RepSerSub 
            Caption         =   "Por Hora"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu RNov_RepSerSub 
            Caption         =   "Por Sentido"
            Enabled         =   0   'False
            Index           =   3
         End
         Begin VB.Menu RNov_RepSerSub 
            Caption         =   "Promedio Móvil"
            Enabled         =   0   'False
            Index           =   4
         End
      End
      Begin VB.Menu RNov_RepTMov 
         Caption         =   "Turno de Móviles"
         Enabled         =   0   'False
         Begin VB.Menu RNov_RepTMoSub 
            Caption         =   "Búsq. por Código"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu RNov_RepTMoSub 
            Caption         =   "Búsq. Por Fecha"
            Enabled         =   0   'False
            Index           =   1
         End
      End
      Begin VB.Menu RNov_RepSu2 
         Caption         =   "Operativos"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu RNov_RepSu2 
         Caption         =   "Total Tareas"
         Enabled         =   0   'False
         Index           =   1
      End
   End
   Begin VB.Menu RNov_Exit 
      Caption         =   "Salir"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu RAcc_Arch 
      Caption         =   "Archivos"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu RAcc_ArcSub 
         Caption         =   "Cía. Seguros"
         Index           =   0
      End
      Begin VB.Menu RAcc_ArcSub 
         Caption         =   "Lugar Traslado"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu RAcc_ArcSub 
         Caption         =   "Marcas"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu RAcc_ArcSub 
         Caption         =   "Patrulleros"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu RAcc_ArcSub 
         Caption         =   "Tipo Vehículos"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu RAcc_ArcSub 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu RAcc_ArcSub 
         Caption         =   "Borrar Ficha"
         Enabled         =   0   'False
         Index           =   6
      End
   End
   Begin VB.Menu RAcc_Fich 
      Caption         =   "Fichas"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu RAcc_FicSub 
         Caption         =   "Ingresar Ficha"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu RAcc_FicSub 
         Caption         =   "Buscar Ficha"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu RAcc_FicSub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu RAcc_FicSub 
         Caption         =   "Nro. Ficha vs Cód. Alfa"
         Index           =   3
      End
   End
   Begin VB.Menu RAcc_Info 
      Caption         =   "Informes"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu RAcc_InfSub 
         Caption         =   "01.Principal"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu RAcc_InfSub 
         Caption         =   "02.Principal Colectora"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu RAcc_InfSub 
         Caption         =   "03.Consulta 2004"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu RAcc_InfSub 
         Caption         =   "04.Peligrosidad por Sector"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu RAcc_InfSub 
         Caption         =   "05.Evaluación"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu RAcc_InfSub 
         Caption         =   "06.Informe Día de Semana"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu RAcc_InfSub 
         Caption         =   "07.Informe Hora Día"
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu RAcc_InfSub 
         Caption         =   "08.Otros Traslados"
         Enabled         =   0   'False
         Index           =   7
      End
      Begin VB.Menu RAcc_InfSub 
         Caption         =   "09.Puntos Negros"
         Enabled         =   0   'False
         Index           =   8
      End
      Begin VB.Menu RAcc_InfSub 
         Caption         =   "10.Informe Personal"
         Enabled         =   0   'False
         Index           =   9
      End
      Begin VB.Menu RAcc_InfSub 
         Caption         =   "11.Peatón Ciclista"
         Enabled         =   0   'False
         Index           =   10
      End
      Begin VB.Menu RAcc_InfSub 
         Caption         =   "12.Tipo Vehículo"
         Enabled         =   0   'False
         Index           =   11
      End
      Begin VB.Menu RAcc_InfSub 
         Caption         =   "13.Total Mensuales"
         Enabled         =   0   'False
         Index           =   12
      End
      Begin VB.Menu RAcc_InfSub 
         Caption         =   "14.Total Discriminado"
         Enabled         =   0   'False
         Index           =   13
      End
      Begin VB.Menu RAcc_InfSub 
         Caption         =   "15.Consulta Parametrizada"
         Enabled         =   0   'False
         Index           =   14
      End
      Begin VB.Menu RAcc_InfSub 
         Caption         =   "16.Demoras de Patrullas"
         Enabled         =   0   'False
         Index           =   15
      End
      Begin VB.Menu RAcc_InfSub 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu RAcc_InfSub 
         Caption         =   "17.Fichas de Accidentes"
         Enabled         =   0   'False
         Index           =   17
      End
      Begin VB.Menu RAcc_InfSub 
         Caption         =   "18.Planillas Mensuales"
         Enabled         =   0   'False
         Index           =   18
      End
      Begin VB.Menu RAcc_InfSub 
         Caption         =   "19.Informe Occovi"
         Enabled         =   0   'False
         Index           =   19
      End
   End
   Begin VB.Menu RAcc_Exit 
      Caption         =   "Salir"
      Visible         =   0   'False
   End
   Begin VB.Menu MEdReg 
      Caption         =   "Registrar"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu MEdRegSol 
         Caption         =   "Solicitudes (Operaciones)"
         Enabled         =   0   'False
      End
      Begin VB.Menu MEdRegSoM 
         Caption         =   "Solicitudes (Mant. Edilicio)"
         Enabled         =   0   'False
      End
      Begin VB.Menu MEdRegRep 
         Caption         =   "Reparaciones"
         Enabled         =   0   'False
      End
      Begin VB.Menu MEdRegAdi 
         Caption         =   "Adicionales"
         Enabled         =   0   'False
      End
      Begin VB.Menu MEdRegVer 
         Caption         =   "Verificacion trabajos"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MEdRep 
      Caption         =   "Reportes"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu MEdRepEst 
         Caption         =   "Estado Solicit."
         Enabled         =   0   'False
         Begin VB.Menu MEdRepEstSec 
            Caption         =   "Operaciones"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu MEdRepEstSec 
            Caption         =   "Mant. Edilicio"
            Enabled         =   0   'False
            Index           =   1
         End
      End
      Begin VB.Menu MEdRepDia 
         Caption         =   "Diarios"
         Enabled         =   0   'False
      End
      Begin VB.Menu MEdRepComp 
         Caption         =   "Completos"
         Enabled         =   0   'False
      End
      Begin VB.Menu MEdRepCompXzona 
         Caption         =   "Completos por zona"
         Enabled         =   0   'False
      End
      Begin VB.Menu MEdRepVer 
         Caption         =   "Verificacion trabajos"
         Enabled         =   0   'False
      End
      Begin VB.Menu MEdRepPendientes 
         Caption         =   "Tareas Pendientes"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MEdSal 
      Caption         =   "Salir"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu MElecReg 
      Caption         =   "Registrar"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu MElecRegSol 
         Caption         =   "Solicitudes"
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRegVal 
         Caption         =   "Validaciones"
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRegSep 
         Caption         =   "-"
      End
      Begin VB.Menu MElecRegNewOT 
         Caption         =   "Nueva O.T."
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRegCloseOT 
         Caption         =   "Cerrar O.T."
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRegAjOT 
         Caption         =   "Ajuste de materiales por O.T."
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRegNewRele 
         Caption         =   "Nuevo Relevamiento"
         Enabled         =   0   'False
         Begin VB.Menu MElecRegNewReleSUB 
            Caption         =   "Nuevo Relevamiento (Genérico)"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu MElecRegNewReleSUB 
            Caption         =   "Nuevo Relevamiento (Columnas)"
            Enabled         =   0   'False
            Index           =   1
         End
      End
      Begin VB.Menu MElecRegCancelarPartes 
         Caption         =   "Cancelar Partes"
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRepReimpOTAbierta 
         Caption         =   "Reimprimir OT abierta"
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRegSepa 
         Caption         =   "-"
      End
      Begin VB.Menu MElecRegNuevoComunicado 
         Caption         =   "Nuevo Comunicado"
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRegSep1 
         Caption         =   "-"
      End
      Begin VB.Menu MElecRegAbastOT 
         Caption         =   "Abastecimiento O.T."
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRegAjAbastOT 
         Caption         =   "Ajuste - Abastecimiento O.T."
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRegIngMateriales 
         Caption         =   "Ingreso de Materiales"
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRegAjIngMateriales 
         Caption         =   "Ajuste - Ingreso de Materiales"
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRegRepoVehiculo 
         Caption         =   "Reposición de Vehículos"
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRegAjusteInven 
         Caption         =   "Ajustes por auditoria"
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRegSep2 
         Caption         =   "-"
      End
      Begin VB.Menu MElecRegRep 
         Caption         =   "Reparaciones"
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRegRel 
         Caption         =   "Relevamientos"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MElecRep 
      Caption         =   "Reportes"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu MElecRepDet 
         Caption         =   "Detalle Solicitudes"
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRepDetNewFormat 
         Caption         =   "Detalle Solicitudes (Nuevo Formato)"
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRepPar 
         Caption         =   "Partes Diarios"
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRepEstadoPartes 
         Caption         =   "Estado de Partes"
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRepPartesLuminarias 
         Caption         =   "Partes de Luminarias"
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRepOrdTrabajo 
         Caption         =   "Ordenes de Trabajo"
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRepEstadoComunicados 
         Caption         =   "Estado de Comunicados"
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRepHistInterColumnas 
         Caption         =   "Histórico de Intervención de Columna"
         Enabled         =   0   'False
      End
      Begin VB.Menu MElecRepStock 
         Caption         =   "Stock"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MElecSal 
      Caption         =   "Salir"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu Inven_Arch 
      Caption         =   "Archivos"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu Inven_ArcSub 
         Caption         =   "Unidades de Medida"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu Inven_ArcSub 
         Caption         =   "Productos"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu Inven_ArcSub 
         Caption         =   "Almacenes"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu Inven_ArcSub 
         Caption         =   "Bodegas"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu Inven_ArcSub 
         Caption         =   "Ubicaciones"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu Inven_ArcSub 
         Caption         =   "Stock Mínimo"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu Inven_ArcSub 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu Inven_ArcSub 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   7
      End
      Begin VB.Menu Inven_ArcSub 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   8
      End
   End
   Begin VB.Menu Inven_Movi 
      Caption         =   "Movimientos"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu Inven_MoviSub 
         Caption         =   "Ingresos - Nueva O.C."
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu Inven_MoviSub 
         Caption         =   "Ingresos - Agregar items a O.C."
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu Inven_MoviSub 
         Caption         =   "Egresos"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu Inven_MoviSub 
         Caption         =   "Ajustes"
         Enabled         =   0   'False
         Index           =   3
      End
   End
   Begin VB.Menu Inven_Repo 
      Caption         =   "Reportes"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu Inven_RepoSub 
         Caption         =   "Stocks"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu Inven_RepoSub 
         Caption         =   "Movimientos"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu Inven_RepoSub 
         Caption         =   "Consumos"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu Inven_RepoSub 
         Caption         =   "Stocks debajo del Mínimo"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu Inven_RepoSub 
         Caption         =   "Egresos por Personal"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu Inven_RepoSub 
         Caption         =   "Ajustes"
         Enabled         =   0   'False
         Index           =   5
      End
   End
   Begin VB.Menu Inven_Exit 
      Caption         =   "Salir"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu CPEEK_IDat 
      Caption         =   "Insertar Datos"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "C. Km 14"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "C. Km 23"
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "C. Km 32"
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "C. Km 36"
         Enabled         =   0   'False
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "C. Km 47"
         Enabled         =   0   'False
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "C. Rta 5"
         Enabled         =   0   'False
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "Conexión Inalámbrica"
         Index           =   7
      End
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "WT_ Km 14"
         Enabled         =   0   'False
         Index           =   9
      End
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "WT_ Km 23"
         Enabled         =   0   'False
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "WT_ Km 32"
         Enabled         =   0   'False
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "WT_ Uruguay(2020)"
         Enabled         =   0   'False
         Index           =   12
      End
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "WT_Testing"
         Enabled         =   0   'False
         Index           =   13
      End
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "WT_Balbin"
         Enabled         =   0   'False
         Index           =   14
      End
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "WT_Belgrano"
         Enabled         =   0   'False
         Index           =   15
      End
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "WT_Boqueron"
         Enabled         =   0   'False
         Index           =   16
      End
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "WT_Escobar"
         Enabled         =   0   'False
         Index           =   17
      End
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "WT_Maschwitz"
         Enabled         =   0   'False
         Index           =   18
      End
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "WT_Pilar"
         Enabled         =   0   'False
         Index           =   19
      End
      Begin VB.Menu CPEEK_IDaCont 
         Caption         =   "WT_Uruguay"
         Enabled         =   0   'False
         Index           =   20
      End
   End
   Begin VB.Menu CPEEK_Info 
      Caption         =   "Informes"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu CPEEK_InfVolu 
         Caption         =   "Volumen"
         Enabled         =   0   'False
         Begin VB.Menu CPEEK_InfVolMatr 
            Caption         =   "Mensual"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu CPEEK_InfVolMatr 
            Caption         =   "Diario"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu CPEEK_InfVolMatr 
            Caption         =   "Diario x Hora"
            Enabled         =   0   'False
            Index           =   2
         End
      End
      Begin VB.Menu CPEEK_InfLong 
         Caption         =   "Longitudes"
         Enabled         =   0   'False
         Begin VB.Menu CPEEK_InfLonMatr 
            Caption         =   "Mensual"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu CPEEK_InfLonMatr 
            Caption         =   "Diario"
            Enabled         =   0   'False
            Index           =   1
         End
      End
      Begin VB.Menu CPEEK_InfVelo 
         Caption         =   "Velocidades"
         Enabled         =   0   'False
         Begin VB.Menu CPEEK_InfVelMatr 
            Caption         =   "Mensual"
            Index           =   0
         End
         Begin VB.Menu CPEEK_InfVelMatr 
            Caption         =   "Diario"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu CPEEK_InfVelMatr 
            Caption         =   "Indicador atascos OEA"
            Index           =   2
         End
      End
      Begin VB.Menu CPEEK_InfVxLo 
         Caption         =   "Veloc. x Long."
         Enabled         =   0   'False
         Begin VB.Menu CPEEK_InfVxLoMatr 
            Caption         =   "Mensual"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu CPEEK_InfVxLoMatr 
            Caption         =   "Diario"
            Enabled         =   0   'False
            Index           =   1
         End
      End
      Begin VB.Menu CPEEK_InfErro 
         Caption         =   "Errores"
         Enabled         =   0   'False
      End
      Begin VB.Menu CPEEK_InfVac 
         Caption         =   "-"
      End
      Begin VB.Menu CPEEK_InfECon 
         Caption         =   "Estado conexiones"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu CPEEK_Exit 
      Caption         =   "Salir"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu ERP_Vacio 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu ERP_Vent 
      Caption         =   "Ventana"
      Visible         =   0   'False
      Begin VB.Menu ERP_VenMini 
         Caption         =   "Minimizar"
      End
      Begin VB.Menu ERP_VenVaci 
         Caption         =   "-"
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "ConsPea"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Reg. Novedades"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Reg. Accidentes"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Guía de Bolsas"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Arqueo Supervisores"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Sist. Violaciones"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Valid"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Pasadas Telepeaje"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Mantenimiento"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Cambio de Clave"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Permisos a Usuarios"
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Nuevos Usuarios"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Gestión de Inventario"
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Contadores PEEK"
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Gestión de Cursos"
         Index           =   14
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Generación de Archivos"
         Index           =   15
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Cambio de Baterias TAG"
         Index           =   16
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Consulta de TAGs"
         Index           =   17
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Balanza Móvil"
         Index           =   18
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "POLAD"
         Index           =   19
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Mantenimiento Edilicio"
         Index           =   20
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Seguridad Informática"
         Index           =   21
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Seg. de Empresas"
         Index           =   22
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Stock Boletos - Tick. Man."
         Index           =   23
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Módulo Gestión IT"
         Enabled         =   0   'False
         Index           =   24
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Telecargas Peaje"
         Index           =   25
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Intranet"
         Index           =   26
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Sistema Gestión Conocimiento"
         Index           =   27
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Gestión CV"
         Index           =   28
         Visible         =   0   'False
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   ""
         Index           =   29
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   ""
         Index           =   30
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   ""
         Index           =   31
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   ""
         Index           =   32
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   ""
         Index           =   33
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   ""
         Index           =   34
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   ""
         Index           =   35
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   ""
         Index           =   36
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   ""
         Index           =   37
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   ""
         Index           =   38
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   ""
         Index           =   39
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   ""
         Index           =   40
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   ""
         Index           =   41
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   ""
         Index           =   42
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   ""
         Index           =   43
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   ""
         Index           =   44
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   ""
         Index           =   45
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   ""
         Index           =   46
      End
      Begin VB.Menu ERP_VenWind 
         Caption         =   "Mantenimiento Electrico"
         Index           =   47
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public mMenuActivo As Integer
Dim mVecSoft(50) As String
Public mUser As String
Public mClave As String
Public Val_mData As Database
Public Mantenim As Database
Public Val_mMenu As Integer
Public mPCname As String
Public mRNovFlag As Boolean
Public mAgregaCpbte As Boolean
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Sub MDIForm_Load()
Dim mi As Integer
MDI.Caption = sMessage
ERP3_frm.Show
mMenuActivo = 999
For mi = 0 To 38
   mVecSoft(mi) = ""
Next
mPCname = PCname
mPCname = Mid(Trim(mPCname), 1, Len(mPCname) - 1)
mRNovFlag = True
End Sub

Public Function PCname() As String
Dim nPC As String
Dim buffer As String
Dim Estado As Long
buffer = String$(255, " ")
Estado = GetComputerName(buffer, 255)
If Estado <> 0 Then
   nPC = Left(buffer, 255)
End If
PCname = nPC
End Function

Public Sub ERP_VenMini_Click()
ERP_VenWind(mMenuActivo).Checked = False
If Screen.ActiveForm.Name <> "MDI" And Screen.ActiveForm.WindowState <> 1 And Screen.ActiveForm.Name <> "ERP1_frm" Then
   mVecSoft(mMenuActivo) = Screen.ActiveForm.Name
   Select Case mMenuActivo
      Case 1
         Select Case mVecSoft(mMenuActivo)
            Case "RNov1a_frm", "RNov1b_frm", "RNov1d_frm"
               RNov1a_frm.WindowState = 1
               RNov1a_frm.Visible = False
               RNov1b_frm.WindowState = 1
               RNov1b_frm.Visible = False
               RNov1d_frm.WindowState = 1
               RNov1d_frm.Visible = False
            Case "RNov2_frm"
               RNov2_frm.WindowState = 1
               RNov2_frm.Visible = False
            Case "RNov3_frm"
               RNov3_frm.WindowState = 1
               RNov3_frm.Visible = False
            Case "RNov5_frm"
               RNov5_frm.WindowState = 1
               RNov5_frm.Visible = False
            Case "RNov6_frm"
               RNov1a_frm.WindowState = 1
               RNov1a_frm.Visible = False
               RNov1b_frm.WindowState = 1
               RNov1b_frm.Visible = False
               RNov1d_frm.WindowState = 1
               RNov1d_frm.Visible = False
               RNov6_frm.WindowState = 1
               RNov6_frm.Visible = False
            Case "RNov7_frm"
               RNov7_frm.WindowState = 1
               RNov7_frm.Visible = False
            Case "RNov8_frm"
               RNov8_frm.WindowState = 1
               RNov8_frm.Visible = False
            Case "RNov9_frm"
               RNov9_frm.WindowState = 1
               RNov9_frm.Visible = False
         End Select
      
      Case Else
         Screen.ActiveForm.WindowState = 1
         Screen.ActiveForm.Visible = False
   End Select
End If
ShowMenu mMenuActivo, False, True
mMenuActivo = 999
ERP_VenMini.Enabled = False
ERP1_frm.Visible = True
MDI.Caption = Left(MDI.Caption, 35)
End Sub

Public Sub ERP_VenWind_Click(Index As Integer)
Dim mi As Integer
If Not ERP_VenWind(Index).Checked Then
   If mMenuActivo <> 999 Then
      mMinimizar
   End If
   ERP_VenWind(Index).Checked = True
   ERP_VenMini.Enabled = True
   mMenuActivo = Index
   ERP1_frm.Visible = False
   If mVecSoft(Index) <> "" Then
      Select Case mVecSoft(Index)
         'SISTEMA DE REGISTRO DE NOVEDADES          Registro de Novedades
         Case "RNov1a_frm", "RNov1b_frm", "RNov1d_frm"
            RNov1a_frm.WindowState = 0
            RNov1a_frm.Visible = True
            RNov1b_frm.WindowState = 0
            RNov1b_frm.Visible = True
            RNov1d_frm.WindowState = 0
            RNov1d_frm.Visible = True
            ERP_Vent.Visible = False
            ERP_Vacio.Visible = False
         Case "RNov2_frm"
            RNov2_frm.WindowState = 0
            RNov2_frm.Visible = True
         Case "RNov5_frm"
            RNov5_frm.WindowState = 0
            RNov5_frm.Visible = True
         Case "RNov6_frm"
            RNov1a_frm.WindowState = 0
            RNov1a_frm.Visible = True
            RNov1b_frm.WindowState = 0
            RNov1b_frm.Visible = True
            RNov1d_frm.WindowState = 0
            RNov1d_frm.Visible = True
            RNov6_frm.WindowState = 0
            RNov6_frm.Visible = True
         Case "RNov7_frm"
            RNov7_frm.WindowState = 0
            RNov7_frm.Visible = True
         Case "RNov8_frm"
            RNov8_frm.WindowState = 0
            RNov8_frm.Visible = True
         Case "RNov9_frm"
            RNov9_frm.WindowState = 0
            RNov9_frm.Visible = True
         'SISTEMA DE REGISTRO DE ACCIDENTES              Registro de Accidentes
         Case "RAcc1beta"
            RAcc1beta.WindowState = 0
            RAcc1beta.Visible = True
         Case "RAcc4_frm"
            RAcc4_frm.WindowState = 0
            RAcc4_frm.Visible = True
         Case "RAcc5_frm"
            RAcc5_frm.WindowState = 0
            RAcc5_frm.Visible = True
         Case "RAcc6_frm"
            RAcc6_frm.WindowState = 0
            RAcc6_frm.Visible = True
         Case "RAcc7_frm"
            RAcc7_frm.WindowState = 0
            RAcc7_frm.Visible = True
         Case "RAcc8_frm"
            RAcc8_frm.WindowState = 0
            RAcc8_frm.Visible = True
         Case "ERP2_frm"
            ERP2_frm.WindowState = 0
            ERP2_frm.Visible = True
         Case "ERP4_frm"
            ERP4_frm.WindowState = 0
            ERP4_frm.Visible = True
         Case "ERP5_frm"
            ERP5_frm.WindowState = 0
            ERP5_frm.Visible = True
         'SISTEMA DE CONTADORES PEEK            Peek
         Case "Pek1_frm"
            Pek1_frm.WindowState = 0
            Pek1_frm.Visible = True
         Case "Peek4_frm"
            Peek4_frm.WindowState = 0
            Peek4_frm.Visible = True
         'Mantenimiento Edilicio
         Case "MEdfrm01"
            MEdfrm01.WindowState = 0
            MEdfrm01.Visible = True
         Case "MEdfrm02"
            MEdfrm02.WindowState = 0
            MEdfrm02.Visible = True
         Case "MEdfrm03"
            MEdfrm03.WindowState = 0
            MEdfrm03.Visible = True
      End Select
      MDI.Caption = Left(MDI.Caption, 35) & Space(10) & " -    Sistema " & ERP_VenWind(Index).Caption
   Else
      ShowMenu Index, True, False
   End If
End If
End Sub

Public Function mFormActive(mSist As Integer)
mVecSoft(mSist) = ""
End Function

Private Function mMinimizar()
If Screen.ActiveForm.WindowState <> 1 And Screen.ActiveForm.Name <> "MDI" And Screen.ActiveForm.Name <> "ERP1_frm" Then
   mVecSoft(mMenuActivo) = Screen.ActiveForm.Name
   Screen.ActiveForm.WindowState = 1
   Screen.ActiveForm.Visible = False
   MDI.Caption = Left(MDI.Caption, 35)
   ShowMenu mMenuActivo, False, True
Else
   If mMenuActivo <> 999 Then
      ShowMenu mMenuActivo, False, True
   End If
End If
End Function

Private Function MostrarForm7(mTable As String, mTittle As String)
ShowMenu 1, False, False
RNov7_frm.mTabla = mTable
RNov7_frm.Label1.Caption = mTittle
RNov7_frm.Show
End Function

Private Sub RNov_ArcSub_Click(Index As Integer)  'ARCHIVOS
Dim mObjRN As New clRNov
Dim Vector(36) As String
Dim i, j, Rep As Integer
Dim CodAlfa As String
Dim mFecha As Date
   Select Case Index
      Case 0  'Emisoras
           RNov2_frm.mTabla = "emisoras"
           RNov2_frm.Label1 = "Tabla Emisoras"
      Case 1  'Origen
           RNov2_frm.mTabla = "origen"
           RNov2_frm.Label1 = "Tabla Origen"
      Case 2  'Otros
           RNov2_frm.mTabla = "otros"
           RNov2_frm.Label1 = "Tabla Otros"
      Case 3  'Patrulleros
           RNov2_frm.mTabla = "patrulleros"
           RNov2_frm.Label1 = "Tabla Patrulleros"
      Case 4  'Rutinas
           RNov2_frm.mTabla = "rutinas"
           RNov2_frm.Label1 = "Tabla Rutinas"
      Case 5  'Servicios
           RNov2_frm.mTabla = "servicios"
           RNov2_frm.Label1 = "Tabla Servicios"
      Case 6  'Tipo Móvil
           RNov2_frm.mTabla = "tipomovil"
           RNov2_frm.Label1 = "Tabla Tipo Movil"
      Case 7  'Tipo Otros
           RNov2_frm.mTabla = "tipootros"
           RNov2_frm.Label1 = "Tabla Tipo Otros"
      Case 8  'Trabajos
      Case 9  'Turnos
           RNov2_frm.mTabla = "turnos"
           RNov2_frm.Label1 = "Tabla Turnos"
      Case 11 'Blanquear Partes
           If MsgBox("¿Está Seguro de Depurar los Números de Partes?", vbYesNo, sMessage) = vbYes Then
              mObjRN.xUpUltimos 0
              MsgBox "Depuración Terminada!!!", vbInformation, sMessage
           End If
      Case 12 'Generar Códigos
           If MsgBox("¿Está Seguro de Generar Códigos Alfanuméricos?", vbYesNo, sMessage) = vbYes Then
              mFecha = Now
              For i = 1 To 10
                 Vector(i) = i - 1
              Next
              For i = 65 To 90
               Vector(i - 54) = Chr(i)
              Next
              mObjRN.xDelCodigos
              MsgBox "Este Proceso Tardará unos Minutos, El Sistema Estará Inhabilitado hasta que Aparezca un Mensaje de Finalización!!!", vbCritical, sMessage
              sMsgEspere Me, "Generando códigos", True
              Randomize
              For j = 1 To 120000
                 CodAlfa = ""
                 For Rep = 1 To 7
                    i = Int((36 * Rnd) + 1)
                    CodAlfa = CodAlfa & Vector(i)
                 Next
                 mObjRN.xInsCodigos CodAlfa
              Next
              sMsgEspere Me, "", False
              MsgBox "Operación Finalizada!! Tiempo Transcurrido = " & DateDiff("s", mFecha, Now) & " Segundos", vbInformation, "Atención!! - RegNov 3.1"
           End If
   End Select
   If Index <> 11 And Index <> 12 Then
      ShowMenu 1, False, False
   End If
   Set mObjRN = Nothing
End Sub

Private Sub RNov_NovRegi_Click() 'Novedad-Registrar
Dim mObjrnov As New clRNov
ShowMenu 1, False, False
If mRNovFlag Then
   mRNovFlag = False
   RNov1a_frm.Visible = True
   RNov1b_frm.Visible = True
   RNov1d_frm.Visible = True
   mObjrnov.xInsertUserRegNov Trim(Right(MDI.mUser, 20)), mPCname
Else
   ShowMenu mMenuActivo, False, True
End If
ERP_Vent.Visible = False
ERP_Vacio.Visible = False
Set mObjrnov = Nothing
End Sub

Private Sub RNov_NovVer_Click() 'Novedades-Ver
ShowMenu 1, False, False
RNov8_frm.Show
End Sub

Private Sub RNov_NovModi_Click() 'Novedades-Modificar
ShowMenu 1, False, False
RNov9_frm.Show
End Sub

Private Sub RNov_NovDepu_Click() 'Novedades-Depurar
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
Dim mFlag As Boolean
Dim mFecha As String
Dim mKm As String
If MsgBox("¿Está seguro que desea Depurar las Novedades?", vbYesNo, sMessage) = vbYes Then
   'pasar las novedades2 a novedades
   mFecha = ""
   sMsgEspere Me, "Depurando... espere.", True
   Set mRec = mObj.oMaxTabla("novedades", "fecha", "")
   If Not mRec.EOF Then
      mFecha = NVL(mRec!Total, "2016-01-01 00:00:00")
   End If
   mRec.Close
   Set mRec = mObj.oTabla("novedades2", "where fecha > '" & Format(mFecha, "yyyy-mm-dd hh:mm:ss") & "' order by fecha")
   If Not mRec.EOF Then
      Do While Not mRec.EOF
         mKm = mRec.Fields(3)
         mFlag = mObj.xInsNov_old(mRec.Fields(0), mRec.Fields(1), mRec.Fields(2), mKm, mRec.Fields(4), mRec.Fields(5), mRec.Fields(6), mRec.Fields(7), mRec.Fields(8), mRec.Fields(9), mRec.Fields(10), mRec.Fields(11), mRec.Fields(12), mRec.Fields(13), mRec.Fields(14), mRec.Fields(15), mRec.Fields(16), mRec.Fields(17), NVL(mRec.Fields(18), ""), mRec.Fields(19), mRec.Fields(20), mRec.Fields(21), mRec.Fields(22))
         mRec.MoveNext
      Loop
   End If
   mRec.Close
   mObj.xDepuraNov
   sMsgEspere Me, "", False
   MsgBox "Depuración Terminada!!!", vbInformation, sMessage
End If
Set mObj = Nothing
End Sub

Private Sub RNov_RepSub_Click(Index As Integer)
Select Case Index
   Case 0
      MostrarForm7 "0", "Consulta Novedades de Móvil"
   Case 1
      MostrarForm7 "1", "Consulta Kilómetros Recorridos por Móvil"
   Case 2
      MostrarForm7 "2", "Consulta Arribos"
   Case 3
      MostrarForm7 "3", "Consulta Códigos de Ambulancia"
   Case 4
      MostrarForm7 "4", "Consulta  de Emisoras"
   Case 5
      MostrarForm7 "5", "Consulta Origen de Novedades"
   Case 6
      MostrarForm7 "16", "Consulta Novedades"
   Case 7
      MostrarForm7 "17", "Retiro de Objetos"
   Case 8
      MostrarForm7 "18", "Reporte Cuadro"
End Select
End Sub

Private Sub RNov_RepRutSub_Click(Index As Integer)
Select Case Index
   Case 0
      MostrarForm7 "6", "Consulta Resumen de Rutinas"
   Case 1
       MostrarForm7 "7", "Consulta Rutinas por Fechas"
   Case 2
       MostrarForm7 "8", "Consulta Detalle de Rutinas"
End Select
End Sub

Private Sub RNov_RepSegSub_Click(Index As Integer)
Select Case Index
  Case 0
      MostrarForm7 "9", "Consulta Seguimiento de Móviles"
  Case 1
      'MsgBox "Consulta Seguimiento Por Móvil y Novedades"
      MostrarForm7 "23", "Consulta Seguimiento Por Móviles (xls)"
  Case 2
      MostrarForm7 "10", "Consulta Seguimiento Por Evento"
End Select
End Sub

Private Sub RNov_RepSerSub_Click(Index As Integer)
Select Case Index
   Case 0
      MostrarForm7 "11", "Consulta Servicios Realizados por Tipo y Móvil"
   Case 1
      MostrarForm7 "12", "Consulta Servicios Realizados Por Progresiva"
   Case 2
      MostrarForm7 "13", "Consulta Servicios Realizados Por Hora"
   Case 3
      MostrarForm7 "14", "Consulta Servicios Realizados Por Sentido"
   Case 4
      MostrarForm7 "15", "Consulta Promedio Móvil"
End Select
End Sub

Private Sub RNov_RepSu2_Click(Index As Integer)
Select Case Index
   Case 0
      MostrarForm7 "21", "Operativos de Móviles"
   Case 1
      MostrarForm7 "22", "Total de Tareas"
End Select
End Sub

Private Sub RNov_RepTMoSub_Click(Index As Integer)
Select Case Index
   Case 0
      MostrarForm7 "19", "Consulta Turnos de Móviles por Código"
   Case 1
      MostrarForm7 "20", "Consulta Turnos de Móviles por Fecha"
End Select
End Sub

Private Sub RNov_Exit_Click()
ShowMenu 1, False, True
ERP1_frm.Visible = True
mExitSist 1
End Sub

Private Sub RAcc_ArcSub_Click(Index As Integer)
ShowMenu 2, False, False
Select Case Index
   Case 0
      RAcc4_frm.mTabla = "CiaSeguros"
      RAcc4_frm.Label1.Caption = "Tabla de Cías de Seguros"
      RAcc4_frm.Show
   Case 1
      RAcc4_frm.mTabla = "LugarTrasl"
      RAcc4_frm.Label1.Caption = "Tabla de Lugar de Traslados"
      RAcc4_frm.Show
   Case 2
      RAcc5_frm.Show
   Case 3
      RNov_ArcSub_Click 3
'      RAcc4_frm.mTabla = "Patrullero"
'      RAcc4_frm.Label1.Caption = "Tabla de Patrulleros"
'      RAcc4_frm.Show
      
'      RNov2_frm.mTabla = "Patrullero"
'      RNov2_frm.Label1 = "Tabla de Patrulleros"
'      RNov2_frm.Show
      
      
      
      
   Case 4
      RAcc4_frm.mTabla = "TipoVehiculo"
      RAcc4_frm.Label1.Caption = "Tabla de Tipo de Vehiculos"
      RAcc4_frm.Show
   
   Case 6 'borrar ficha
      RAcc13.Show
End Select
End Sub

Private Sub RAcc_ArcCSeg_Click()
ShowMenu 2, False, False
End Sub

Private Sub RAcc_ArcLTra_Click()
RAcc4_frm.mTabla = "LugarTrasl"
RAcc4_frm.Label1.Caption = "Tabla de Lugar de Traslados"
RAcc4_frm.Show
End Sub

Private Sub RAcc_ArcMarc_Click()
ShowMenu 2, False, False
End Sub

Private Sub RAcc_ArcPatr_Click()
ShowMenu 2, False, False
End Sub

Private Sub RAcc_ArcTVeh_Click()
ShowMenu 2, False, False
End Sub

Private Sub RAcc_FicSub_Click(Index As Integer)
   ShowMenu 2, False, False
   Select Case Index
      Case 0 'Ingresar Ficha
         RAcc1beta.Show
      Case 1 'Buscar Ficha
         RAcc1beta.Show
         RAcc1beta.sInitModif
      Case 3  'Nro. Ficha Vs Código alfanumérico.
         RAcc12.Show
   End Select
End Sub

Private Sub RAcc_InfSub_Click(Index As Integer)
 ShowMenu 2, False, False
 Select Case Index
   Case 0
       RAcc6_frm.mReporte = "Principal"
   Case 1
       RAcc6_frm.mReporte = "Principal Colectora"
   Case 2
       RAcc6_frm.mReporte = "Consulta_2004"
   Case 3
       RAcc6_frm.mReporte = "Peligrosidad por Sector"
   Case 4
       RAcc6_frm.mReporte = "Evaluación"
   Case 5
       RAcc6_frm.mReporte = "Día de Semana"
   Case 6
       RAcc6_frm.mReporte = "Hora Día"
   Case 7
       RAcc6_frm.mReporte = "Otros Traslados"
   Case 8
       RAcc6_frm.mReporte = "Puntos Negros"
   Case 9
       RAcc6_frm.mReporte = "Informe Personal"
   Case 10
       RAcc6_frm.mReporte = "Informe Personal"
   Case 11
       RAcc6_frm.mReporte = "Tipo Vehículo"
   Case 12
       RAcc6_frm.mReporte = "Total Mensuales"
   Case 13
       RAcc6_frm.mReporte = "Total Discriminado"
   Case 14
       RAcc7_frm.Show
   Case 15
       RAcc6_frm.mReporte = "Demoras de Patrullas"
   Case 17
      RAcc10_frm.Show
   Case 18
      RAcc11_frm.Show
   Case 19
      RAcc14.Show
 End Select
 If Index <> 14 And Index <> 17 And Index <> 18 And Index <> 19 Then
    RAcc6_frm.Show
 End If
End Sub

Private Sub RAcc_Inf01_Click()
   ShowMenu 2, False, False
   RAcc6_frm.mReporte = "Principal"
   RAcc6_frm.Show
End Sub

Private Sub RAcc_Inf02_Click()
   ShowMenu 2, False, False
   RAcc6_frm.mReporte = "Principal Colectora"
   RAcc6_frm.Show
End Sub

Private Sub RAcc_Inf03_Click()
   ShowMenu 2, False, False
   RAcc6_frm.mReporte = "Consulta_2004"
   RAcc6_frm.Show
End Sub

Private Sub RAcc_Inf04_Click()
   ShowMenu 2, False, False
   RAcc6_frm.mReporte = "Peligrosidad por Sector"
   RAcc6_frm.Show
End Sub

Private Sub RAcc_Inf05_Click()
   ShowMenu 2, False, False
   RAcc6_frm.mReporte = "Evaluación"
   RAcc6_frm.Show
End Sub

Private Sub RAcc_Inf06_Click()
   ShowMenu 2, False, False
   RAcc6_frm.mReporte = "Día de Semana"
   RAcc6_frm.Show
End Sub

Private Sub RAcc_Inf07_Click()
   ShowMenu 2, False, False
   RAcc6_frm.mReporte = "Hora Día"
   RAcc6_frm.Show
End Sub

Private Sub RAcc_Inf08_Click()
   ShowMenu 2, False, False
   RAcc6_frm.mReporte = "Otros Traslados"
   RAcc6_frm.Show
End Sub

Private Sub RAcc_Inf09_Click()
   ShowMenu 2, False, False
   RAcc6_frm.mReporte = "Puntos Negros"
   RAcc6_frm.Show
End Sub

Private Sub RAcc_Inf10_Click()
   ShowMenu 2, False, False
   RAcc6_frm.mReporte = "Informe Personal"
   RAcc6_frm.Show
End Sub

Private Sub RAcc_Inf11_Click()
   ShowMenu 2, False, False
   RAcc6_frm.mReporte = "Peatón Ciclista"
   RAcc6_frm.Show
End Sub

Private Sub RAcc_Inf12_Click()
   ShowMenu 2, False, False
   RAcc6_frm.mReporte = "Tipo Vehículo"
   RAcc6_frm.Show
End Sub

Private Sub RAcc_Inf13_Click()
   ShowMenu 2, False, False
   RAcc6_frm.mReporte = "Total Mensuales"
   RAcc6_frm.Show
End Sub

Private Sub RAcc_Inf14_Click()
   ShowMenu 2, False, False
   RAcc6_frm.mReporte = "Total Discriminado"
   RAcc6_frm.Show
End Sub

Private Sub RAcc_Inf15_Click()
   ShowMenu 2, False, False
   RAcc7_frm.Show
End Sub

Private Sub RAcc_Inf16_Click()
ShowMenu 2, False, False
RAcc6_frm.mReporte = "Demoras de Patrullas"
RAcc6_frm.Show
End Sub

Private Sub RAcc_Exit_Click()
ShowMenu 2, False, True
ERP1_frm.Visible = True
ERP1_frm.Top = 0
ERP1_frm.Left = 0
mExitSist 2
End Sub

Private Sub CPEEK_IDaCont_Click(Index As Integer) 'Ingreso de Datos
Dim mObj As New clPeek
   Select Case Index
      Case 0
        mObj.PasarDatos "Km14"
      Case 1
        mObj.PasarDatos "Km23"
      Case 2
        mObj.PasarDatos "Km32"
      Case 3
        mObj.PasarDatos "Km36"
      Case 4
        mObj.PasarDatos "Km47"
      Case 5
        mObj.PasarDatos "Rta5"
      Case 7
         Peek_3.Show
         ShowMenu 13, False, False
      Case 9   ' WT_km 14
        mObj.fImportWT "km14"
      Case 10  ' WT_km 23
        mObj.fImportWT "km23"
      Case 11  ' WT_km 32
        mObj.fImportWT "km32"
      Case 12  ' WT_km 32
        'mObj.fImportWT "km36"
        mObj.fImportWT "U2ug"
      Case 13  ' WT_Test
        mObj.fImportWT "Test"
      Case 14  ' WT_Balb
        mObj.fImportWT "Balb"
      Case 15  ' WT_Belg
        mObj.fImportWT "Belg"
      Case 16  ' WT_Boqu
        mObj.fImportWT "Boqu"
      Case 17  ' WT_Esco
        mObj.fImportWT "Esco"
      Case 18  ' WT_Masc
        mObj.fImportWT "Masc"
      Case 19  ' WT_Pi46
        mObj.fImportWT "Pi46"
      Case 20  ' WT_Urug
        mObj.fImportWT "Urug"
   End Select
   Set mObj = Nothing
End Sub

Private Sub CPEEK_InfVolMatr_Click(Index As Integer)
   Pek1_frm.Show
   Pek1_frm.Command1(0).Tag = Index
   Pek1_frm.Command1(1).Tag = "0"
   Select Case Index
      Case 0
         Pek1_frm.Label2.Caption = "Volumétrico Mensual"   'Informes-Volumétrico-Mensual
      Case 1
         Pek1_frm.Label2.Caption = "Volumétrico Diario"   'Informes-Volumétrico-Mensual
      Case 2
          Pek1_frm.Label2.Caption = "Volumétrico Diario x Hora"   'Informes-Volumétrico-Diario x Hora
   End Select
   ShowMenu 13, False, False
End Sub

Private Sub CPEEK_InfLonMatr_Click(Index As Integer) 'Informes-Longitudes-Mensual / Informes-Longitudes-Diario
   Pek1_frm.Show
   Pek1_frm.Command1(0).Tag = Index
   Pek1_frm.Command1(1).Tag = "1"
   Select Case Index
      Case 0
         Pek1_frm.Label2.Caption = "LONGITUDES Mensual"
      Case 1
         Pek1_frm.Label2.Caption = "LONGITUDES Diario"
   End Select
   ShowMenu 13, False, False
End Sub

'Private Sub CPEEK_InfVelMatr_Click(Index As Integer)
'   Pek1_frm.Show
'   Pek1_frm.Command1(0).Tag = Index
'   Pek1_frm.Command1(1).Tag = "2"
'   If Index = 0 Then
'      Pek1_frm.Label2.Caption = "VELOCIDADES Mensual"
'   Else
'      Pek1_frm.Label2.Caption = "VELOCIDADES Diario"
'   End If
'   ShowMenu 13, False, False
'End Sub

Private Sub CPEEK_InfVelMatr_Click(Index As Integer)
   Pek1_frm.Show
   Pek1_frm.Command1(0).Tag = Index
   Pek1_frm.Command1(1).Tag = "2"
'   If Index = 0 Then
'      Pek1_frm.Label2.Caption = "VELOCIDADES Mensual"
'   Else
'      Pek1_frm.Label2.Caption = "VELOCIDADES Diario"
'   End If
   
   
   Select Case Index
      Case 0
         Pek1_frm.Label2.Caption = "VELOCIDADES Mensual"
      Case 1
         Pek1_frm.Label2.Caption = "VELOCIDADES Diario"
      Case 2
         Pek1_frm.Label2.Caption = "Indidcador atascos OEA"
   End Select
      
   ShowMenu 13, False, False
End Sub


Private Sub CPEEK_InfVxLoMatr_Click(Index As Integer)
   Pek1_frm.Show
   Pek1_frm.Command1(0).Tag = Index
   Pek1_frm.Command1(1).Tag = "4"
   If Index = 0 Then
      Pek1_frm.Label2.Caption = "VELOCIDADES x LONGITUDES Mensual"
   Else
      Pek1_frm.Label2.Caption = "VELOCIDADES x LONGITUDES Diario"
   End If
   ShowMenu 13, False, False
End Sub

Private Sub CPEEK_InfErro_Click() 'Informes-Errores
   Pek1_frm.Show
   Pek1_frm.Command1(1).Tag = "3"
   Pek1_frm.Label2.Caption = "Informe de Errores en Ingreso de Datos"
   Pek1_frm.Combo1(0).AddItem "TODOS"
   ShowMenu 13, False, False
End Sub

Private Sub CPEEK_InfECon_Click()
Peek4_frm.Show
ShowMenu 13, False, False
End Sub

Private Sub CPEEK_Exit_Click()
ShowMenu 13, False, True
ERP1_frm.Visible = True
mExitSist 13
End Sub

Private Sub MEdRegAdi_Click()
ShowMenu 20, False, True
MEdfrm05.Show
End Sub

Private Sub MEdRegRep_Click()
ShowMenu 20, False, True
MEdfrm02.Show
End Sub

Private Sub MEdRegSol_Click()
ShowMenu 20, False, True
MEdfrm01.Show
End Sub

Private Sub MEdRegSoM_Click()
ShowMenu 20, False, True
MEdfrm04.Show
End Sub

Private Sub MEdRegVer_Click()
ShowMenu 20, False, True
MEdfrm06.Show
End Sub

Private Sub MEdRepComp_Click()
ShowMenu 20, False, True
MEdfrm03.mSector = 3
MEdfrm03.Show
End Sub


Private Sub MEdRepCompXzona_Click()
ShowMenu 20, False, True
MEdfrm03.mSector = 5
MEdfrm03.Show
End Sub

Private Sub MEdRepDia_Click()
ShowMenu 20, False, True
MEdfrm03.mSector = 2
MEdfrm03.Show
End Sub


Private Sub MEdRepPendientes_Click()
   MEdfrm03.GenXLS_PartesPendientes
End Sub

Private Sub MEdRepEstSec_Click(Index As Integer)
ShowMenu 20, False, True
MEdfrm03.mSector = Index
MEdfrm03.Show
End Sub

Private Sub MEdRepVer_Click()
ShowMenu 20, False, True
MEdfrm03.mSector = 4
MEdfrm03.Show
End Sub

Private Sub MEdSal_Click()
ShowMenu 20, False, True
ERP1_frm.Visible = True
mExitSist 20
End Sub

Private Sub MElecRegRel_Click()
ShowMenu 47, False, True
MantElect05.Show
End Sub

Private Sub MElecRegRep_Click()
ShowMenu 47, False, True
MantElect02.Show
End Sub

Private Sub MElecRegSol_Click()
ShowMenu 47, False, True
MantElect01.Show
End Sub

Private Sub MElecRegVal_Click()
ShowMenu 47, False, True
MantElect03.Show
End Sub

Private Sub MElecRegNewOT_Click()
ShowMenu 47, False, True
MantElect06.Show
End Sub

Private Sub MElecRegCloseOT_Click()
ShowMenu 47, False, True
MantElect08.Show
End Sub
Private Sub MElecRegAjOT_Click()
ShowMenu 47, False, True
MantElect11.Show
End Sub

Private Sub MElecRegNuevoComunicado_Click()
ShowMenu 47, False, True
MantElect14.Show
End Sub

Private Sub MElecRegNewReleSUB_Click(Index As Integer)
If Index = 0 Then
   ShowMenu 47, False, True
   MantElect16.Show
Else
   ShowMenu 47, False, True
   MantElect18.Show
End If
End Sub

Private Sub MElecRepEstadoPartes_Click()
ShowMenu 47, False, True
MantElect19.Show
End Sub

Private Sub MElecRepOrdTrabajo_Click()
ShowMenu 47, False, True
MantElect20.Show
End Sub

Private Sub MElecRepEstadoComunicados_Click()
ShowMenu 47, False, True
MantElect21.Show
End Sub

Private Sub MElecRepPartesLuminarias_Click()
ShowMenu 47, False, True
MantElect22.Show
End Sub

Private Sub MElecRepHistInterColumnas_Click()
ShowMenu 47, False, True
MantElect23.Show
End Sub

Private Sub MElecRepReimpOTAbierta_Click()
ShowMenu 47, False, True
MantElect24.Show
End Sub

Private Sub MElecRegAjusteInven_Click()
   ShowMenu 47, False, True
   Inven017_frm.Show
End Sub

Private Sub MElecRegCancelarPartes_Click()
   ShowMenu 47, False, True
   MantElect17.Show
End Sub

Private Sub MElecRegAbastOT_Click()
ShowMenu 47, False, True
MantElect09.Show
End Sub

Private Sub MElecRegAjAbastOT_Click()
ShowMenu 47, False, True
MantElect12.Show
End Sub

Private Sub MElecRegIngMateriales_Click()
ShowMenu 47, False, True
MantElect10.Show
End Sub

Private Sub MElecRegAjIngMateriales_Click()
ShowMenu 47, False, True
MantElect13.Show
End Sub

Private Sub MElecRegRepoVehiculo_Click()
ShowMenu 47, False, True
MantElect15.Show
End Sub

Private Sub MElecRepDet_Click()
ShowMenu 47, False, True
MantElect04.mReporte = 0
MantElect04.Show
End Sub

Private Sub MElecRepPar_Click()
ShowMenu 47, False, True
MantElect04.mReporte = 1
MantElect04.Show

End Sub

Private Sub MElecRepDetNewFormat_Click()
ShowMenu 47, False, True
MantElect04.mReporte = 2
MantElect04.Show
End Sub

Private Sub MElecRepStock_Click()
ShowMenu 47, False, True
Inven013_frm.Show
End Sub




Private Sub MElecSal_Click()
ShowMenu 47, False, True
ERP1_frm.Visible = True
mExitSist 47
End Sub

Private Sub Inven_ArcSub_Click(Index As Integer)
Dim mObjInv As New clInven
   Select Case Index
      Case 0  'Unidades de Medida
           Inven2_frm.mTabla = "UnidadMedida"
           Inven2_frm.Label1 = "Tabla Unidades de medida"
      Case 1  'Productos
           Inven3_frm.Show
      Case 2  'Almacen
           Inven2_frm.mTabla = "Almacenes"
           Inven2_frm.Label1 = "Tabla Almacenes"
      Case 3  'Bodega
           Inven7_frm.Show
      Case 4 'Ubicacion
           Inven8_frm.Show
      Case 5 'Stock Mínimo
           Inven9_frm.Show
   End Select
   
   ShowMenu 12, False, True
  Set mObjInv = Nothing
End Sub

Private Sub Inven_MoviSub_Click(Index As Integer)
   
   ShowMenu 12, False, True

   Select Case Index
      Case 0
         Inven011_frm.Show
      Case 1
         'Inven6_frm.Show
         Inven014_frm.Show
      Case 2
         'Inven6_frm.Show
         Inven010_frm.Show
      Case 3
         Inven017_frm.Show
   End Select
End Sub

Private Sub Inven_RepoSub_Click(Index As Integer)
   ShowMenu 12, False, True
   
   Select Case Index
      Case 0
         Inven013_frm.Show
      Case 1
         Inven012_frm.Caption = "Movimientos"
         Inven012_frm.mReporte = "Movimientos"
         Inven012_frm.Show
      Case 2
         Inven015_frm.Show
      Case 3
         Inven013_frm.StockDebajoDelMinimo
      Case 4
         Inven016_frm.Show
      Case 5
         Inven012_frm.Caption = "Ajustes"
         Inven012_frm.mReporte = "Ajustes"
         Inven012_frm.Show
   End Select
End Sub

Private Sub Inven_Exit_Click()
ShowMenu 12, False, True
ERP1_frm.Visible = True
mExitSist 12
End Sub




