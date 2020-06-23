VERSION 5.00
Begin VB.Form ERP6_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1065
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7860
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   7860
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   420
      Width           =   6495
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   6900
      Stretch         =   -1  'True
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   1080
      Left            =   5
      TabIndex        =   0
      Top             =   5
      Width           =   7835
   End
End
Attribute VB_Name = "ERP6_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public pTexto As String

Private Sub Form_Load()
   Me.Height = 1140
   Me.Width = 7890
   sAlinearForm Me
   Label2.Caption = pTexto
   Image1.Picture = LoadPicture(App.Path & "\ERP\Imagenes\Relojarena2.jpg")
End Sub
