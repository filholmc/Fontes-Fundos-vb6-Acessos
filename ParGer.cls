VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clssParGer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Function Consultar(ByVal vstrCodigo As String) As Recordset
       Screen.MousePointer = vbHourglass

          Set Consultar = New Recordset
              Consultar.Open _
 _
             "SELECT * FROM TBParGer WHERE Codigo = '" & vstrCodigo & "'", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Sub Alterar(ByVal vstrCodigo As String, ByVal vstrCteudo As String)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "UPDATE TBParGer SET Cteudo = '" & vstrCteudo & "' WHERE Codigo = '" & vstrCodigo & "'"

       Screen.MousePointer = vbDefault
End Sub
