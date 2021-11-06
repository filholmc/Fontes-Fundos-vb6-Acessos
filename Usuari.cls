VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clssUsuari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Function Consultar(ByVal vintNumero As Integer) As Recordset
       Screen.MousePointer = vbHourglass

          Set Consultar = New Recordset
              Consultar.Open _
 _
             "SELECT * FROM TBUsuari WHERE Numero = " & vintNumero, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarTodosPorNumero() As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarTodosPorNumero = New Recordset
              ConsultarTodosPorNumero.Open _
 _
             "SELECT * FROM TBUsuari ORDER BY Numero", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Sub Incluir(ByVal vintNumero As Integer, ByVal vstrSenhas As String, ByVal vbytTemLog As Byte, _
                                                                            ByVal vbytStatus As Byte, _
                                                                            ByVal vstrNomUsu As String, _
                                                                            ByVal vintCodAge As Integer, _
                                                                            ByVal vdatDatVal As Date, _
                                                                            ByVal vstrFuncao As String, _
                                                                            ByVal vstrE_Mail As String)
       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "INSERT INTO TBUsuari       (Numero,              Senhas,              TemLog,             Status,              NomUsu,              CodAge,                     DatVal,                              Funcao,               E_Mail)     " & _
                                         "    VALUES (" & vintNumero & ", '" & vstrSenhas & "', " & vbytTemLog & ", " & vbytStatus & ", '" & vstrNomUsu & "', " & vintCodAge & ", '" & Format(vdatDatVal, "dd/mm/yyyy") & "', '" & vstrFuncao & "', '" & vstrE_Mail & "')"
       Screen.MousePointer = vbDefault
End Sub
Public Sub Alterar(ByVal vintNumero As Integer, ByVal vbytTemLog As Byte, ByVal vbytStatus As String, _
                                                                          ByVal vstrNomUsu As String, _
                                                                          ByVal vintCodAge As String, _
                                                                          ByVal vdatDatVal As String, _
                                                                          ByVal vstrFuncao As String, _
                                                                          ByVal vstrE_Mail As String)
       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "UPDATE TBUsuari SET TemLog =  " & vbytTemLog & ",  Status =  " & vbytStatus & _
                                         ",           NomUsu = '" & vstrNomUsu & "', CodAge =  " & vintCodAge & _
                                         ",           DatVal = '" & Format(vdatDatVal, "dd/mm/yyyy") & _
                                         "',          Funcao = '" & vstrFuncao & "', E_Mail = '" & vstrE_Mail & _
                                         "'     WHERE Numero =  " & vintNumero
       Screen.MousePointer = vbDefault
End Sub
Public Sub AlterarSenha(ByVal vintNumero As Integer, ByVal vstrSenhas As String)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "UPDATE TBUsuari SET Senhas = '" & vstrSenhas & "' WHERE Numero = " & vintNumero

       Screen.MousePointer = vbDefault
End Sub
Public Sub Excluir(ByVal vintNumero As Integer)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "DELETE FROM TBUsuari WHERE Numero = " & vintNumero

       Screen.MousePointer = vbDefault
End Sub
