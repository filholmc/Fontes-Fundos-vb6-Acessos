VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clssAceMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Function Consultar(ByVal vintNumUsu As Integer, ByVal vbytNumMod As Byte) As Recordset
       Screen.MousePointer = vbHourglass

          Set Consultar = New Recordset
              Consultar.Open _
 _
             "SELECT * FROM TBAceMod WHERE NumUsu = " & vintNumUsu & " AND NumMod = " & vbytNumMod, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarModulosDeUmUsuario(ByVal vintNumUsu As Integer) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarModulosDeUmUsuario = New Recordset
              ConsultarModulosDeUmUsuario.Open _
 _
             "SELECT Numero, Descri FROM TBAceMod                " & _
                             "      JOIN TBModulo                " & _
                             "        ON  (NumMod =     Numero)  " & _
                             "     WHERE   NumUsu = " & vintNumUsu & _
                             "  ORDER BY   NumMod   ", _
 _
              gDBCFundos, adOpenStatic, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarUsuariosDeUmModulo(ByVal vbytNumMod As Byte) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarUsuariosDeUmModulo = New Recordset
              ConsultarUsuariosDeUmModulo.Open _
 _
             "SELECT Numero, NomUsu, CodAge FROM TBAceMod                " & _
                                     "      JOIN TBUsuari                " & _
                                     "        ON  (NumUsu =     Numero)  " & _
                                     "     WHERE   NumMod = " & vbytNumMod & _
                                     "  ORDER BY   NomUsu ", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function Ausente(ByVal vintNumUsu As Integer, ByVal vbytNumMod As Byte) As Boolean
       Dim lRSt As Recordset

       Screen.MousePointer = vbHourglass

          Set lRSt = New Recordset
              lRSt.Open _
 _
             "SELECT NumUsu FROM TBAceMod WHERE NumUsu = " & vintNumUsu & " AND NumMod = " & vbytNumMod, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

              Ausente = IIf(lRSt.EOF, True, False)

              lRSt.Close

       Screen.MousePointer = vbDefault
End Function
Public Function ModuloAcessado(ByVal vbytNumMod As Byte) As Boolean
       Dim lRSt As Recordset

       Screen.MousePointer = vbHourglass

          Set lRSt = New Recordset
              lRSt.Open _
 _
             "SELECT TOP 1 NumMod FROM TBAceMod WHERE NumMod = " & vbytNumMod, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

              ModuloAcessado = IIf(lRSt.EOF, False, True)

              lRSt.Close

       Screen.MousePointer = vbDefault
End Function
Public Function UsuarioAcessa(ByVal vintNumUsu As Integer) As Boolean
       Dim lRSt As Recordset

       Screen.MousePointer = vbHourglass

          Set lRSt = New Recordset
              lRSt.Open _
 _
             "SELECT TOP 1 NumUsu FROM TBAceMod WHERE NumUsu = " & vintNumUsu, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

              UsuarioAcessa = IIf(lRSt.EOF, False, True)

              lRSt.Close

       Screen.MousePointer = vbDefault
End Function
Public Sub Incluir(ByVal vintNumUsu As Integer, ByVal vbytNumMod As Byte)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "INSERT INTO TBAceMod       (NumUsu,             NumMod)    " & _
                                         "    VALUES (" & vintNumUsu & ", " & vbytNumMod & ")"
       Screen.MousePointer = vbDefault
End Sub
Public Sub Excluir(ByVal vintNumUsu As Integer, ByVal vbytNumMod As Byte)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "DELETE FROM TBAceMod WHERE NumUsu = " & vintNumUsu & " AND NumMod = " & vbytNumMod

       Screen.MousePointer = vbDefault
End Sub
Public Sub ExcluirTodosDeUmUsuario(ByVal vintNumUsu As Integer)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "DELETE FROM TBAceMod WHERE NumUsu = " & vintNumUsu

       Screen.MousePointer = vbDefault
End Sub
