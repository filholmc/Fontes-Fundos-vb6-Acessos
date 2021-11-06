VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clssDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Function ConsultarDataForm(ByVal vbytNumMod As Byte, ByVal vintNumFor As Integer, _
                                                            ByVal vbytNumFun As Byte, _
                                                            ByVal vstrOrdChv As String) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarDataForm = New Recordset
              ConsultarDataForm.Open _
 _
             "SELECT TOP 1 DatBas FROM TBDiario WHERE NumMod = " & vbytNumMod & _
                                       "          AND NumFor = " & vintNumFor & _
                                       "          AND NumFun = " & vbytNumFun & _
                                       "     ORDER BY DatBas   " & vstrOrdChv, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarDataUsuario(ByVal vintNumUsu As Integer, ByVal vstrOrdChv As String) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarDataUsuario = New Recordset
              ConsultarDataUsuario.Open _
 _
             "SELECT TOP 1 DatBas FROM TBDiario WHERE NumUsu = " & vintNumUsu & _
                                       "     ORDER BY DatBas   " & vstrOrdChv, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarDeUmForm(ByVal vbytNumMod As Byte, ByVal vintNumFor As Integer, _
                                                            ByVal vbytNumFun As Byte, _
                                                            ByVal vdatDatIni As Date, _
                                                            ByVal vdatDatFim As Date) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarDeUmForm = New Recordset
              ConsultarDeUmForm.Open _
 _
             "SELECT * FROM TBDiario WHERE NumMod  =  " & vbytNumMod & _
                            "          AND NumFor  =  " & vintNumFor & _
                            "          AND NumFun  =  " & vbytNumFun & _
                            "          AND DatBas >= '" & Format(vdatDatIni, "dd/mm/yyyy") & _
                            "'         AND DatBas <= '" & Format(vdatDatFim, "dd/mm/yyyy") & _
                            "'    ORDER BY DatBas  ", _
 _
              gDBCFundos, adOpenStatic, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarDeUmUsuario(ByVal vintNumUsu As String, ByVal vdatDatIni As Date, _
                                                                 ByVal vdatDatFim As Date) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarDeUmUsuario = New Recordset
              ConsultarDeUmUsuario.Open _
 _
             "SELECT * FROM TBDiario WHERE NumUsu  =  " & vintNumUsu & _
                            "          AND DatBas >= '" & Format(vdatDatIni, "dd/mm/yyyy") & _
                            "'         AND DatBas <= '" & Format(vdatDatFim, "dd/mm/yyyy") & _
                            "'    ORDER BY DatBas  ", _
 _
              gDBCFundos, adOpenStatic, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarFormes(ByVal vbytNumMod As Byte) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarFormes = New Recordset
              ConsultarFormes.Open _
 _
             "SELECT DISTINCT NumFor, Descri FROM TBDiario                       " & _
                                      "      JOIN TBFormes                       " & _
                                      "        ON         (NumFor =     Numero)  " & _
                                      "     WHERE TBDiario.NumMod = " & vbytNumMod & _
                                      "  ORDER BY          Descri   ", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarFundos(ByVal vbytNumMod As Byte, ByVal vintNumFor As Integer) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarFundos = New Recordset
              ConsultarFundos.Open _
 _
             "SELECT DISTINCT NumFun, Codigo FROM TBDiario                " & _
                                      "      JOIN TBFundos                " & _
                                      "        ON  (NumFun =     Numero)  " & _
                                      "     WHERE   NumMod = " & vbytNumMod & _
                                      "       AND   NumFor = " & vintNumFor & _
                                      "  ORDER BY   Codigo   ", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarModulos() As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarModulos = New Recordset
              ConsultarModulos.Open _
 _
             "SELECT DISTINCT NumMod FROM TBDiario ORDER BY NumMod", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarUsuarios() As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarUsuarios = New Recordset
              ConsultarUsuarios.Open _
 _
             "SELECT DISTINCT NumUsu FROM TBDiario ORDER BY NumUsu", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Sub Incluir(ByVal vbooForLog As Boolean, ByVal vintNumUsu As Integer, ByVal vstrNomCmp As String, _
                                                                             ByVal vbytNumMod As Byte, _
                                                                             ByVal vintNumFor As Integer, _
                                                                             ByVal vstrFuncao As String, _
                                                                             ByVal vbytNumFun As Byte, _
                                                                             ByVal vstrChaves As String, _
                                                                             ByVal vstrCteudo As String)
       On Error Resume Next

       If (Not gbooUsuLog And Not vbooForLog) Then Exit Sub

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "INSERT INTO TBDiario       (NumUsu,             DatBas,                       NomCmp,              NumMod,             NumFor,              Funcao,              NumFun,              Chaves,               Cteudo)     " & _
                                         "    VALUES (" & vintNumUsu & ", '" & rgfDataDoServidor & "', '" & vstrNomCmp & "', " & vbytNumMod & ", " & vintNumFor & ", '" & vstrFuncao & "', " & vbytNumFun & ", '" & vstrChaves & "', '" & vstrCteudo & "')"
       Screen.MousePointer = vbDefault
End Sub
