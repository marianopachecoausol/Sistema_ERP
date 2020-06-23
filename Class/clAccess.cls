VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private mData As Database

Private Sub Class_Terminate()
Set mData = Nothing
End Sub

Public Function mBorrarAuxi(ByVal pBase As String, ByVal pTable As String)
Dim mI As Integer
Dim mBorrar As Boolean
Set mData = OpenDatabase(App.Path & pBase & ".mdb")
mBorrar = False
For mI = 0 To mData.TableDefs.Count - 1
   If mData.TableDefs(mI).Name = pTable Then
      mBorrar = True
   End If
Next
If mBorrar Then
   mData.Execute "DROP TABLE " & pTable & ""
End If
mData.Close
End Function
