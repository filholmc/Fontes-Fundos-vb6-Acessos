VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clssFundos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Function Consultar(ByVal vbytNumero As Byte) As Recordset
       Screen.MousePointer = vbHourglass

          Set Consultar = New Recordset
              Consultar.Open _
 _
             "SELECT * FROM TBFundos WHERE Numero = " & vbytNumero, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarTodos() As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarTodos = New Recordset
              ConsultarTodos.Open _
 _
             "SELECT * FROM TBFundos ORDER BY Codigo", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function CodigoCadastrado(ByVal vstrCodigo As String) As Boolean
       Dim lRSt As Recordset

       Screen.MousePointer = vbHourglass

          Set lRSt = New Recordset
              lRSt.Open _
 _
             "SELECT Codigo FROM TBFundos WHERE Codigo = '" & vstrCodigo & "'", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

              CodigoCadastrado = IIf(lRSt.EOF, False, True)

              lRSt.Close

       Screen.MousePointer = vbDefault
End Function
Public Sub Incluir(ByVal vbytNumero As Byte, ByVal vstrCodigo As String)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "INSERT INTO TBFundos       (Numero,              Codigo)     " & _
                                         "    VALUES (" & vbytNumero & ", '" & vstrCodigo & "')"
       Screen.MousePointer = vbDefault
End Sub
Public Sub Alterar(ByVal vbytNumero As Byte, ByVal vstrCodigo As String)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "UPDATE TBFundos SET Codigo = '" & vstrCodigo & "' WHERE Numero = " & vbytNumero

       Screen.MousePointer = vbDefault
End Sub
Public Sub Excluir(ByVal vbytNumero As Byte)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "DELETE FROM TBFundos WHERE Numero = " & vbytNumero

       Screen.MousePointer = vbDefault
End Sub
