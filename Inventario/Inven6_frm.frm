VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Inven6_frm 
   Caption         =   "Form2"
   ClientHeight    =   13920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   26550
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   13920
   ScaleWidth      =   26550
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4095
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   5400
      Width           =   9855
      Begin VB.TextBox Text2 
         Height          =   975
         Left            =   600
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   600
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4095
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   9855
      Begin VB.TextBox Text1 
         Height          =   855
         Left            =   480
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   720
         Width           =   3975
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   26535
      _ExtentX        =   46805
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Orden de Trabajo"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Inven6_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TabStrip1_Click()
   Dim i As Integer
   
    i = TabStrip1.SelectedItem.Index
    'Mostrar el contenedor que corresponda
   
   Select Case i
      
      Case 1
        Frame1(0).Visible = True
        Frame1(1).Visible = False
      Case 2
        Frame1(0).Visible = False
        Frame1(1).Visible = True
   End Select

End Sub
