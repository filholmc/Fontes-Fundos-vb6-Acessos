VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clssLogado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Function Consultar(ByVal vstrNomCmp As String) As Recordset
       Screen.MousePointer = vbHourglass

          Set Consultar = New Recordset
              Consultar.Open _
 _
             "SELECT * FROM TBLogado WHERE NomCmp = '" & vstrNomCmp & "'", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Sub Incluir(ByVal vstrNomCmp As String, ByVal vintNumUsu As Integer, ByVal vintCodAge As Integer, _
                                                                            ByVal vstrCodFun As String, _
                                                                            ByVal vstrModulo As String, _
                                                                            ByVal vbytModAtv As Byte)
       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "INSERT INTO TBLogado        (NomCmp,              NumUsu,             CodAge,              CodFun,               Modulo,              ModAtv)    " & _
                                         "    VALUES ('" & vstrNomCmp & "', " & vintNumUsu & ", " & vintCodAge & ", '" & vstrCodFun & "', '" & vstrModulo & "', " & vbytModAtv & ")"
       Screen.MousePointer = vbDefault
End Sub
Public Sub Alterar(ByVal vstrNomCmp As String, ByVal vintNumUsu As Integer, ByVal vbytModAtv As Byte)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "UPDATE TBLogado SET NumUsu =  " & vintNumUsu & _
                                         ",           Modatv =  " & vbytModAtv & _
                                         "      WHERE NomCmp = '" & vstrNomCmp & "'"
       Screen.MousePointer = vbDefault
End Sub
Public Sub Excluir(ByVal vstrNomCmp As String)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "DELETE FROM TBLogado WHERE NomCmp = '" & vstrNomCmp & "'"

       Screen.MousePointer = vbDefault
End Sub
