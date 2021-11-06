VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clssModulo"
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
             "SELECT * FROM TBModulo WHERE Numero = " & vbytNumero, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarTodos() As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarTodos = New Recordset
              ConsultarTodos.Open _
 _
             "SELECT * FROM TBModulo ORDER BY Numero", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function DescricaoCadastrada(ByVal vstrDescri As String) As Boolean
       Dim lRSt As Recordset

       Screen.MousePointer = vbHourglass

          Set lRSt = New Recordset
              lRSt.Open _
 _
             "SELECT Descri FROM TBModulo WHERE Descri = '" & vstrDescri & "'", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

              DescricaoCadastrada = IIf(lRSt.EOF, False, True)

              lRSt.Close

       Screen.MousePointer = vbDefault
End Function
Public Sub Incluir(ByVal vbytNumero As Byte, ByVal vstrDescri As String)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "INSERT INTO TBModulo       (Numero,              Descri)     " & _
                                         "    VALUES (" & vbytNumero & ", '" & vstrDescri & "')"
       Screen.MousePointer = vbDefault
End Sub
Public Sub Alterar(ByVal vbytNumero As Byte, ByVal vstrDescri As String)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "UPDATE TBModulo SET Descri = '" & vstrDescri & "' WHERE Numero = " & vbytNumero

       Screen.MousePointer = vbDefault
End Sub
Public Function Excluir(ByVal vbytNumero As Byte) As Errors

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "DELETE FROM TBModulo WHERE Numero = " & vbytNumero

       Screen.MousePointer = vbDefault
End Function
