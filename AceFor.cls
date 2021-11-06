VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clssAceFor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Function Consultar(ByVal vintNumUsu As Integer, ByVal vintNumFor As Integer) As Recordset
       Screen.MousePointer = vbHourglass

          Set Consultar = New Recordset
              Consultar.Open _
 _
             "SELECT * FROM TBAceFor WHERE NumUsu = " & vintNumUsu & " AND NumFor = " & vintNumFor, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarFormsDeUmUsuarioPorModulo(ByVal vintNumUsu As Integer, ByVal vbytNumMod As Byte) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarFormsDeUmUsuarioPorModulo = New Recordset
              ConsultarFormsDeUmUsuarioPorModulo.Open _
 _
             "SELECT Numero, NomFor, Descri FROM TBAceFor                " & _
                                     "      JOIN TBFormes                " & _
                                     "        ON  (NumFor =     Numero)  " & _
                                     "     WHERE   NumUsu = " & vintNumUsu & _
                                     "       AND   NumMod = " & vbytNumMod & _
                                     "  ORDER BY   NomFor   ", _
 _
              gDBCFundos, adOpenStatic, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarUsuariosDeUmForm(ByVal vintNumFor As Integer) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarUsuariosDeUmForm = New Recordset
              ConsultarUsuariosDeUmForm.Open _
 _
             "SELECT Numero, NomUsu, CodAge FROM TBAceFor                " & _
                                     "      JOIN TBUsuari                " & _
                                     "        ON  (NumUsu =     Numero)  " & _
                                     "     WHERE   NumFor = " & vintNumFor & _
                                     "  ORDER BY   NomUsu   ", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function Ausente(ByVal vintNumUsu As Integer, ByVal vintNumFor As Integer) As Boolean
       Dim lRSt As Recordset

       Screen.MousePointer = vbHourglass

          Set lRSt = New Recordset
              lRSt.Open _
 _
             "SELECT NumUsu FROM TBAceFor WHERE NumUsu = " & vintNumUsu & " AND NumFor = " & vintNumFor, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

              Ausente = IIf(lRSt.EOF, True, False)

              lRSt.Close

       Screen.MousePointer = vbDefault
End Function
Public Function FormAcessado(ByVal vintNumFor As Integer) As Boolean
       Dim lRSt As Recordset

       Screen.MousePointer = vbHourglass

          Set lRSt = New Recordset
              lRSt.Open _
 _
             "SELECT TOP 1 NumFor FROM TBAceFor WHERE NumFor = " & vintNumFor, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

              FormAcessado = IIf(lRSt.EOF, False, True)

              lRSt.Close

       Screen.MousePointer = vbDefault
End Function
Public Function UsuarioAcessa(ByVal vintNumUsu As Integer) As Boolean
       Dim lRSt As Recordset

       Screen.MousePointer = vbHourglass

          Set lRSt = New Recordset
              lRSt.Open _
 _
             "SELECT TOP 1 NumUsu FROM TBAceFor WHERE NumUsu = " & vintNumUsu, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

              UsuarioAcessa = IIf(lRSt.EOF, False, True)

              lRSt.Close

       Screen.MousePointer = vbDefault
End Function
Public Sub Incluir(ByVal vintNumUsu As Integer, ByVal vintNumFor As Integer)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "INSERT INTO TBAceFor       (NumUsu,             NumFor)    " & _
                                         "    VALUES (" & vintNumUsu & ", " & vintNumFor & ")"
       Screen.MousePointer = vbDefault
End Sub
Public Sub Excluir(ByVal vintNumUsu As Integer, ByVal vintNumFor As Integer)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "DELETE FROM TBAceFor WHERE NumUsu = " & vintNumUsu & " AND NumFor = " & vintNumFor

       Screen.MousePointer = vbDefault
End Sub
Public Sub ExcluirTodosDeUmUsuario(ByVal vintNumUsu As Integer)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "DELETE FROM TBAceFor WHERE NumUsu = " & vintNumUsu

       Screen.MousePointer = vbDefault
End Sub
