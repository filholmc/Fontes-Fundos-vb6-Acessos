VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clssFormes"
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
             "SELECT * FROM TBFormes WHERE Numero = " & vintNumero, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarAjuda(ByVal vbytNumMod As Byte, ByVal vbytNumAju As Byte) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarAjuda = New Recordset
              ConsultarAjuda.Open _
 _
             "SELECT Numero FROM TBFormes WHERE NumMod = " & vbytNumMod & _
                                 "          AND NumAju = " & vbytNumAju, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarFormsDeUmModulo(ByVal vbytNumMod As Byte) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarFormsDeUmModulo = New Recordset
              ConsultarFormsDeUmModulo.Open _
 _
             "SELECT * FROM TBFormes WHERE NumMod = " & vbytNumMod & " ORDER BY NomFor", _
 _
              gDBCFundos, adOpenStatic, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarFormsDeUmModuloAcessaveis(ByVal vbytNumMod As Byte) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarFormsDeUmModuloAcessaveis = New Recordset
              ConsultarFormsDeUmModuloAcessaveis.Open _
 _
             "SELECT * FROM TBFormes WHERE NumMod =   " & vbytNumMod & _
                            "          AND SemAce = 0 " & _
                            "     ORDER BY NomFor     ", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarFormsDeUmModuloComBotoes(ByVal vbytNumMod As Byte) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarFormsDeUmModuloComBotoes = New Recordset
              ConsultarFormsDeUmModuloComBotoes.Open _
 _
             "SELECT DISTINCT NumFor, NomFor FROM  TBFormes                       " & _
                                      "      JOIN  TBBotoes                       " & _
                                      "        ON (TBFormes.Numero =     NumFor)  " & _
                                      "     WHERE  TBFormes.NumMod = " & vbytNumMod & _
                                      "  ORDER BY           NomFor   ", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarFormsDeUmModuloComBotoesDeUmUsuario(ByVal vbytNumMod As Byte, _
                                                             ByVal vintNumUsu As Integer) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarFormsDeUmModuloComBotoesDeUmUsuario = New Recordset
              ConsultarFormsDeUmModuloComBotoesDeUmUsuario.Open _
 _
             "SELECT DISTINCT TBFormes.Numero, NomFor FROM  TBFormes                          " & _
                                               "      JOIN  TBBotoes                          " & _
                                               "        ON (TBFormes.Numero = TBBotoes.NumFor)" & _
                                               "      JOIN  TBAceFor                          " & _
                                               "        ON (TBFormes.Numero = TBAceFor.NumFor)" & _
                                               "     WHERE  TBFormes.NumMod =  " & vbytNumMod & _
                                               "       AND           NumUsu =  " & vintNumUsu & _
                                               "  ORDER BY           NomFor    ", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarNome(ByVal vbytNumMod As Byte, ByVal vstrNomFor As String) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarNome = New Recordset
              ConsultarNome.Open _
 _
             "SELECT Numero FROM TBFormes WHERE NumMod =  " & vbytNumMod & _
                                 "          AND NomFor = '" & vstrNomFor & "'", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarParaAjuda(ByVal vbytNumMod As Byte, ByVal vstrNomFor As String) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarParaAjuda = New Recordset
              ConsultarParaAjuda.Open _
 _
             "SELECT Numero        , NumAju FROM TBFormes WHERE NumMod =  " & vbytNumMod & _
                                                 "          AND NomFor = '" & vstrNomFor & "'", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function PertenceModulo(ByVal vbytNumMod As Byte) As Boolean
       Dim lRSt As Recordset

       Screen.MousePointer = vbHourglass

          Set lRSt = New Recordset
              lRSt.Open _
 _
             "SELECT TOP 1 Numero FROM TBFormes WHERE NumMod = " & vbytNumMod, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

              PertenceModulo = IIf(lRSt.EOF, False, True)

              lRSt.Close

       Screen.MousePointer = vbDefault
End Function
Public Sub Incluir(ByVal vbytNumMod As Byte, ByVal vintNumero As Integer, ByVal vstrNomFor As String, _
                                                                          ByVal vstrDescri As String, _
                                                                          ByVal vbytSemAce As Byte, _
                                                                          ByVal vbytTemLog As Byte, _
                                                                          ByVal vbytNumAju As Byte)
       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "INSERT INTO TBFormes       (NumMod,             Numero,              NomFor,               Descri,              TemLog,             SemAce,             NumAju)    " & _
                                         "    VALUES (" & vbytNumMod & ", " & vintNumero & ", '" & vstrNomFor & "', '" & vstrDescri & "', " & vbytTemLog & ", " & vbytSemAce & ", " & vbytNumAju & ")"
       Screen.MousePointer = vbDefault
End Sub
Public Sub Alterar(ByVal vbytNumMod As Byte, ByVal vintNumero As Integer, ByVal vstrNomFor As String, _
                                                                          ByVal vstrDescri As String, _
                                                                          ByVal vbytSemAce As Byte, _
                                                                          ByVal vbytTemLog As Byte, _
                                                                          ByVal vbytNumAju As Byte)
       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "UPDATE TBFormes SET NumMod =  " & vbytNumMod & ",  NomFor = '" & vstrNomFor & _
                                         "',          Descri = '" & vstrDescri & "', SemAce =  " & vbytSemAce & _
                                         ",           TemLog =  " & vbytTemLog & ",  NumAju =  " & vbytNumAju & _
                                         "      WHERE Numero =  " & vintNumero
       Screen.MousePointer = vbDefault
End Sub
Public Sub Excluir(ByVal vintNumero As Integer)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "DELETE FROM TBFormes WHERE Numero = " & vintNumero

       Screen.MousePointer = vbDefault
End Sub
