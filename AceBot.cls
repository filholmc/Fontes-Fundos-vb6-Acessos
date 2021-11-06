VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clssAceBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Function Consultar(ByVal vintNumUsu As Integer, ByVal vintNumBot As Integer) As Recordset
       Screen.MousePointer = vbHourglass

          Set Consultar = New Recordset
              Consultar.Open _
 _
             "SELECT * FROM TBAceBot WHERE NumUsu = " & vintNumUsu & " AND NumBot = " & vintNumBot, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarBotoesDeUmUsuarioPorModulo(ByVal vintNumUsu As Integer, ByVal vbytNumMod As Byte) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarBotoesDeUmUsuarioPorModulo = New Recordset
              ConsultarBotoesDeUmUsuarioPorModulo.Open _
 _
             "SELECT NumBot FROM TBAceBot                " & _
                     "      JOIN TBBotoes                " & _
                     "        ON  (NumBot =     Numero)  " & _
                     "     WHERE   NumUsu = " & vintNumUsu & _
                     "       AND   NumMod = " & vbytNumMod, _
 _
              gDBCFundos, adOpenStatic, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarBotoesDeUmUsuarioPorModuloAndForm(ByVal vintNumUsu As Integer, _
                                                           ByVal vbytNumMod As Byte, _
                                                           ByVal vintNumFor As Integer) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarBotoesDeUmUsuarioPorModuloAndForm = New Recordset
              ConsultarBotoesDeUmUsuarioPorModuloAndForm.Open _
 _
             "SELECT Numero, NomBot, Descri FROM TBAceBot                " & _
                                     "      JOIN TBBotoes                " & _
                                     "        ON  (NumBot =     Numero)  " & _
                                     "     WHERE   NumUsu = " & vintNumUsu & _
                                     "       AND   NumMod = " & vbytNumMod & _
                                     "       AND   NumFor = " & vintNumFor, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarUsuariosDeUmBotao(ByVal vintNumBot As Integer) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarUsuariosDeUmBotao = New Recordset
              ConsultarUsuariosDeUmBotao.Open _
 _
             "SELECT Numero, NomUsu, CodAge FROM TBAceBot                " & _
                                     "      JOIN TBUsuari                " & _
                                     "        ON  (NumUsu =     Numero)  " & _
                                     "     WHERE   NumBot = " & vintNumBot & _
                                     "  ORDER BY   NomUsu   ", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function Ausente(ByVal vintNumUsu As Integer, ByVal vintNumBot As Integer) As Boolean
       Dim lRSt As Recordset

       Screen.MousePointer = vbHourglass

          Set lRSt = New Recordset
              lRSt.Open _
 _
             "SELECT NumUsu FROM TBAceBot WHERE NumUsu = " & vintNumUsu & " AND NumBot = " & vintNumBot, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

              Ausente = IIf(lRSt.EOF, True, False)

              lRSt.Close

       Screen.MousePointer = vbDefault
End Function
Public Function BotaoAcessado(ByVal vintNumBot As Integer) As Boolean
       Dim lRSt As Recordset

       Screen.MousePointer = vbHourglass

          Set lRSt = New Recordset
              lRSt.Open _
 _
             "SELECT TOP 1 NumBot FROM TBAceBot WHERE NumBot = " & vintNumBot, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

              BotaoAcessado = IIf(lRSt.EOF, False, True)

              lRSt.Close

       Screen.MousePointer = vbDefault
End Function
Public Function UsuarioAcessa(ByVal vintNumUsu As Integer) As Boolean
       Dim lRSt As Recordset

       Screen.MousePointer = vbHourglass

          Set lRSt = New Recordset
              lRSt.Open _
 _
             "SELECT TOP 1 NumUsu FROM TBAceBot WHERE NumUsu = " & vintNumUsu, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

              UsuarioAcessa = IIf(lRSt.EOF, False, True)

              lRSt.Close

       Screen.MousePointer = vbDefault
End Function
Public Sub Incluir(ByVal vintNumUsu As Integer, ByVal vintNumBot As Integer)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "INSERT INTO TBAceBot       (NumUsu,             NumBot)    " & _
                                         "    VALUES (" & vintNumUsu & ", " & vintNumBot & ")"
       Screen.MousePointer = vbDefault
End Sub
Public Sub Excluir(ByVal vintNumUsu As Integer, ByVal vintNumBot As Integer)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "DELETE FROM TBAceBot WHERE NumUsu = " & vintNumUsu & " AND NumBot = " & vintNumBot

       Screen.MousePointer = vbDefault
End Sub
Public Sub ExcluirTodosDeUmUsuario(ByVal vintNumUsu As Integer)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "DELETE FROM TBAceBot WHERE NumUsu = " & vintNumUsu

       Screen.MousePointer = vbDefault
End Sub
