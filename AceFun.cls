VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clssAceFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Function Consultar(ByVal vintNumUsu As Integer, ByVal vbytNumFun As Byte) As Recordset
       Screen.MousePointer = vbHourglass

          Set Consultar = New Recordset
              Consultar.Open _
 _
             "SELECT * FROM TBAceFun WHERE NumUsu = " & vintNumUsu & " AND NumFun = " & vbytNumFun, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarFundosDeUmUsuario(ByVal vintNumUsu As Integer) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarFundosDeUmUsuario = New Recordset
              ConsultarFundosDeUmUsuario.Open _
 _
             "SELECT Numero, Codigo FROM TBAceFun                " & _
                             "      JOIN TBFundos                " & _
                             "        ON  (NumFun =     Numero)  " & _
                             "     WHERE   NumUsu = " & vintNumUsu & _
                             "  ORDER BY   Codigo   ", _
 _
              gDBCFundos, adOpenStatic, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarUsuariosDeUmFundo(ByVal vbytNumFun As Byte) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarUsuariosDeUmFundo = New Recordset
              ConsultarUsuariosDeUmFundo.Open _
 _
             "SELECT Numero, NomUsu, CodAge FROM TBAceFun                " & _
                                     "      JOIN TBUsuari                " & _
                                     "        ON  (NumUsu =     Numero)  " & _
                                     "     WHERE   NumFun = " & vbytNumFun & _
                                     "  ORDER BY   NomUsu ", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function Ausente(ByVal vintNumUsu As Integer, ByVal vbytNumFun As Byte) As Boolean
       Dim lRSt As Recordset

       Screen.MousePointer = vbHourglass

          Set lRSt = New Recordset
              lRSt.Open _
 _
             "SELECT * FROM TBAceFun WHERE NumUsu = " & vintNumUsu & " AND NumFun = " & vbytNumFun, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

              Ausente = IIf(lRSt.EOF, True, False)

              lRSt.Close

       Screen.MousePointer = vbDefault
End Function
Public Function FundoAcessado(ByVal vbytNumFun As Byte) As Boolean
       Dim lRSt As Recordset

       Screen.MousePointer = vbHourglass

          Set lRSt = New Recordset
              lRSt.Open _
 _
             "SELECT TOP 1 NumFun FROM TBAceFun WHERE NumFun = " & vbytNumFun, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

              FundoAcessado = IIf(lRSt.EOF, False, True)

              lRSt.Close

       Screen.MousePointer = vbDefault
End Function
Public Function UsuarioAcessa(ByVal vintNumUsu As Integer) As Boolean
       Dim lRSt As Recordset

       Screen.MousePointer = vbHourglass

          Set lRSt = New Recordset
              lRSt.Open _
 _
             "SELECT TOP 1 NumUsu FROM TBAceFun WHERE NumUsu = " & vintNumUsu, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

              UsuarioAcessa = IIf(lRSt.EOF, False, True)

              lRSt.Close

       Screen.MousePointer = vbDefault
End Function
Public Sub Incluir(ByVal vintNumUsu As Integer, ByVal vbytNumFun As Byte)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "INSERT INTO TBAceFun       (NumUsu,             NumFun)    " & _
                                         "    VALUES (" & vintNumUsu & ", " & vbytNumFun & ")"
       Screen.MousePointer = vbDefault
End Sub
Public Sub Excluir(ByVal vintNumUsu As Integer, ByVal vbytNumFun As Byte)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "DELETE FROM TBAceFun WHERE NumUsu = " & vintNumUsu & " AND NumFun = " & vbytNumFun

       Screen.MousePointer = vbDefault
End Sub
Public Sub ExcluirTodosDeUmUsuario(ByVal vintNumUsu As Integer)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "DELETE FROM TBAceFun WHERE NumUsu = " & vintNumUsu

       Screen.MousePointer = vbDefault
End Sub
