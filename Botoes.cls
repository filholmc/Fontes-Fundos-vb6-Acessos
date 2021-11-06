VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clssBotoes"
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
             "SELECT * FROM TBBotoes WHERE Numero = " & vintNumero, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarBotoesDeUmForm(ByVal vbytNumMod As Byte, ByVal vintNumFor As Integer) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarBotoesDeUmForm = New Recordset
              ConsultarBotoesDeUmForm.Open _
 _
             "SELECT * FROM TBBotoes WHERE NumMod = " & vbytNumMod & _
                            "          AND NumFor = " & vintNumFor & _
                            "     ORDER BY NomBot   ", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarBotoesDeUmModulo(ByVal vbytNumMod As Byte) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarBotoesDeUmModulo = New Recordset
              ConsultarBotoesDeUmModulo.Open _
 _
             "SELECT Numero FROM TBBotoes WHERE NumMod = " & vbytNumMod, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function ConsultarNome(ByVal vbytNumMod As Byte, ByVal vintNumFor As Integer, _
                                                        ByVal vstrNomBot As String) As Recordset
       Screen.MousePointer = vbHourglass

          Set ConsultarNome = New Recordset
              ConsultarNome.Open _
 _
             "SELECT Numero FROM TBBotoes WHERE NumMod =  " & vbytNumMod & _
                                 "          AND NumFor =  " & vintNumFor & _
                                 "          AND NomBot = '" & vstrNomBot & "'", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

       Screen.MousePointer = vbDefault
End Function
Public Function PertenceForm(ByVal vintNumFor As Integer) As Boolean
       Dim lRSt As Recordset

       Screen.MousePointer = vbHourglass

          Set lRSt = New Recordset
              lRSt.Open _
 _
             "SELECT TOP 1 Numero FROM TBBotoes WHERE NumFor = " & vintNumFor, _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

              PertenceForm = IIf(lRSt.EOF, False, True)

              lRSt.Close

       Screen.MousePointer = vbDefault
End Function
Public Sub Incluir(ByVal vbytNumMod As Byte, ByVal vintNumFor As Integer, ByVal vintNumero As Integer, _
                                                                          ByVal vstrNomBot As String, _
                                                                          ByVal vstrDescri As String)
       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "INSERT INTO TBBotoes       (NumMod,             NumFor,             Numero,              NomBot,               Descri)     " & _
                                         "    VALUES (" & vbytNumMod & ", " & vintNumFor & ", " & vintNumero & ", '" & vstrNomBot & "', '" & vstrDescri & "')"
       Screen.MousePointer = vbDefault
End Sub
Public Sub Alterar(ByVal vbytNumMod As Byte, ByVal vintNumFor As Integer, ByVal vintNumero As Integer, _
                                                                          ByVal vstrNomBot As String, _
                                                                          ByVal vstrDescri As String)
       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "UPDATE TBBotoes SET NumMod =  " & vbytNumMod & ",  NumFor =  " & vintNumFor & _
                                         ",           NomBot = '" & vstrNomBot & "', Descri = '" & vstrDescri & _
                                         "'     WHERE Numero =  " & vintNumero
       Screen.MousePointer = vbDefault
End Sub
Public Sub Excluir(ByVal vintNumero As Integer)

       On Error Resume Next

       Screen.MousePointer = vbHourglass

              gDBCFundos.Errors.Clear
              gDBCFundos.Execute "DELETE FROM TBBotoes WHERE Numero = " & vintNumero

       Screen.MousePointer = vbDefault
End Sub
