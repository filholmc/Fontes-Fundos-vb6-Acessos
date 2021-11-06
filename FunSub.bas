Attribute VB_Name = "moduFunSub"
Option Explicit

Public Declare _
       Function rgfNomeDoComputador Lib "kernel32" Alias "GetComputerNameA" (ByVal vstrNomCmp As String, _
                                                                                   vlngTamNom As Long) As Long
Public Function rgfArquivoExiste(ByVal vstrNomArq As String) As Boolean
       Dim lintNumArq As Integer

       On Error GoTo NaoExiste

       Screen.MousePointer = vbHourglass

              lintNumArq = FreeFile

         Open vstrNomArq For Input As #lintNumArq

       Close #lintNumArq

       Screen.MousePointer = vbDefault

       rgfArquivoExiste = True

       Exit Function

NaoExiste:
       rgfArquivoExiste = False

       Screen.MousePointer = vbDefault
End Function
Public Function rgfConectado() As Boolean
       On Error GoTo Erro_DB

       If (Not gDBCFundos Is Nothing) Then
           rgfConectado = True
           Exit Function
       End If

       Set gDBCFundos = New Connection

      With gDBCFundos
          .Provider = "SQLOLEDB"
          .ConnectionString = "Data Source     = " & gstrServBD & ";" & _
                              "Initial Catalog = " & gstrNomeBD & ";" & _
                              "User ID         = " & gstrUsuaBD & ";" & _
                              "Password        = " & gstrSenhBD
          .Open
          .Execute "SET LANGUAGE Portuguese"
       End With

       rgfConectado = True

       Exit Function

Erro_DB:
       rgfConectado = False
End Function
Public Function rgfConfiguracoesRegionaisNaoOk() As Boolean
       Dim lstrMensag As String

           lstrMensag = ""

       If (Format(DateSerial(1966, 6, 21), "short date") <> "21/06/1966") Then
           lstrMensag = "- Na ficha Data, no item Data abreviada:" & vbCr & _
                        "  - Estilo data abreviada = dd/MM/aaaa" & vbCr & _
                        "  - Separador de Data = /" & vbCr & vbCr
       End If

       If (Right(Format(123456.1234, "currency"), 10) <> "123.456,12") Then
           lstrMensag = lstrMensag & "- Na ficha Moeda:" & vbCr & _
                                     "  - Posição do símbolo da moeda = ¤1,1" & vbCr & _
                                     "  - Símbolo decimal = ," & vbCr & _
                                     "  - No. de dígitos decimais = 2" & vbCr & vbCr
       End If

       If (Format(123456.1234, "standard") <> "123.456,12") Then
           lstrMensag = lstrMensag & "- Na ficha Número:" & vbCr & _
                                     "  - Símbolo decimal = ," & vbCr & _
                                     "  - No. de dígitos decimais = 2" & vbCr & _
                                     "  - Símbolo de agrupamento de dígitos = ." & vbCr & vbCr
       End If

       If (lstrMensag = "") Then
           rgfConfiguracoesRegionaisNaoOk = False
       Else
           rgfConfiguracoesRegionaisNaoOk = True
           lstrMensag = "Para evitar problemas na execução, " & vbCr & _
                        "você deve fazer as seguintes alterações " & vbCr & _
                        "no Painel de Controle em Configurações Regionais:" & vbCr & vbCr & lstrMensag & _
                        "Este programa será encerrado agora!" & vbCr & vbCr & _
                        "Faça as Alterações recomendadas e depois execute-o novamente!"
           rgfMsgBox lstrMensag, MsgErr
       End If
End Function
Public Function rgfDataDoServidor() As Date
       Dim lRSt As Recordset

       Screen.MousePointer = vbHourglass

          Set lRSt = New Recordset
              lRSt.Open _
 _
             "SELECT GetDate() AS DatSrv", _
 _
              gDBCFundos, adOpenForwardOnly, adLockReadOnly

              rgfDataDoServidor = lRSt!DatSrv

              lRSt.Close

       Screen.MousePointer = vbDefault
End Function
Public Function rgfSemEdicao(ByVal vstrValors As String) As String
       rgfSemEdicao = Replace(vstrValors, ".", "")
End Function
Public Function rgfSenhaCp(ByVal vstrSenhas As String) As String
       Dim lintCtdor1 As Integer

       Dim lintCtdor2 As Integer

       Dim lstrChvCrp As String * 10

           rgfSenhaCp = ""
           lstrChvCrp = Chr(1) + Chr(2) + Chr(3) + Chr(4) + Chr(5) + Chr(6) + Chr(7) + Chr(8) + Chr(9) + Chr(10)

       For lintCtdor1 = 1 To Len(vstrSenhas)
           lintCtdor2 = lintCtdor2 Mod 10 + 1
           rgfSenhaCp = rgfSenhaCp + Chr(Asc(Mid$(vstrSenhas, lintCtdor1, 1)) Xor _
                                         Asc(Mid$(lstrChvCrp, lintCtdor2, 1)))
       Next
End Function
Public Sub Main()
       Dim lRStLogado As Recordset

       Dim lRStUsuari As Recordset

       Dim lstrPthAAj As String

       Dim lstrPthExe As String

       If (rgfConfiguracoesRegionaisNaoOk) Then End

           rgsLerUsuarioDoDataBase

           lstrPthExe = gstrPthExe & "Acessos.exe"
           lstrPthAAj = gstrPthAAj & "Acessos.chm"

       If (gstrPthAtu <> "" And Dir(gstrPthAtu) <> "") And _
          (lstrPthExe <> "" And Dir(lstrPthExe) <> "") Then
       If (FileDateTime(gstrPthAtu) > FileDateTime(lstrPthExe)) Then
           Shell App.Path & "\Upgrade.exe Acessos " & gstrPthAtu & " " & lstrPthExe, vbNormalFocus
           End
       End If
       End If

       If (lstrPthAAj <> "" And Dir(lstrPthAAj) <> "") And _
          (gstrPthAju <> "" And Dir(gstrPthAju) <> "") Then
       If (FileDateTime(lstrPthAAj) > FileDateTime(gstrPthAju)) Then
           Shell App.Path & "\Upgrade.exe Acessos " & lstrPthAAj & " " & gstrPthAju & " " & lstrPthExe, vbNormalFocus
           End
       End If
       End If

       If (Not rgfConectado) Then
           rgfMsgBox "Erro na Abertura do Banco de Dados '" & gstrNomeBD & "', Usuário '" & gstrUsuaBD & "'" & vbCr & vbCr & _
                     "- Cheque as Conexões de Rede;" & vbCr & vbCr & _
                     "- Verifique o funcionamento do Servidor '" & gstrServBD & "'", MsgErr
       End
       End If

       Set gclsUsuari = New clssUsuari
       Set gclsFormes = New clssFormes
       Set gclsAceMod = New clssAceMod
       Set gclsLogado = New clssLogado
       Set gclsDiario = New clssDiario
       Set gclsInsAdm = New clssInsAdm

           rgsConsultarAdministradora
'          __________________________________________________________________________ Obtém o Nome do Computador na Rede

           gstrNomCmp = Space(30)

           rgfNomeDoComputador gstrNomCmp, 30

           gstrNomCmp = Trim(gstrNomCmp)
           gstrNomCmp = Mid((gstrNomCmp), 1, Len(gstrNomCmp) - 1)
'          ______________________________________________________________ Ativa cópia ou alterna para uma já em execução

       If (App.PrevInstance) Then
           rgfMsgBox "Este Módulo já está em execução!  " & vbCr & vbCr & _
                     "Ative-o na Barra de Tarefas.      " & vbCr & vbCr & _
                     "Esta nova execução será encerrada.", MsgInf

       Set gclsUsuari = Nothing
       Set gclsFormes = Nothing
       Set gclsAceMod = Nothing
       Set gclsLogado = Nothing
       Set gclsDiario = Nothing
       Set gclsInsAdm = Nothing
       End
       End If
'          _______________________________________________________________________ Exibe a Tela de Logar ou entra direto

       Set lRStLogado = gclsLogado.Consultar(gstrNomCmp)

       If (Command = "") Then

           formAcesso.Show 1

       If (gbooCancel) Then End
       Else
           gintUsuLog = lRStLogado!NumUsu
           gintAgeLog = lRStLogado!CodAge
           gbytModAtv = lRStLogado!ModAtv + 1

       Set lRStUsuari = gclsUsuari.Consultar(gintUsuLog)

           gstrSenhas = lRStUsuari!Senhas
'          gbooUsuLog = IIf( _
'                       lRStUsuari!TemLog, 1, 0)
           gstrNomUsu = lRStUsuari!NomUsu

           lRStUsuari.Close

           gclsLogado.Alterar gstrNomCmp, gintUsuLog, gbytModAtv

       If (gDBCFundos.Errors.Count > 0) Then rgfMsgBox "Erro no Ajuste dos Módulos ativos", MsgErr
       End If
       lRStLogado.Close
       formMDIAce.Show
End Sub
Public Sub rgsCarregarFormesDeUmModuloAcessaveis(ByVal vbytNumMod As Byte, ByRef vcboFormes As ComboBox, _
                                                                           ByVal vbooItem00 As Boolean)
       Dim lRStFormes As Recordset

       Set lRStFormes = gclsFormes.ConsultarFormsDeUmModuloAcessaveis(vbytNumMod)

           vcboFormes.Clear

       If (vbooItem00) Then _
           vcboFormes.AddItem "Todos - 0"
       Do _
           While (Not lRStFormes.EOF)
           vcboFormes.AddItem lRStFormes!NomFor & " - " & lRStFormes!Numero
           lRStFormes.MoveNext
       Loop
       lRStFormes.Close
End Sub
Public Sub rgsCarregarFormesDeUmModuloComBotoes(ByVal vbytNumMod As Byte, ByRef vcboFormes As ComboBox)
       Dim lRStFormes As Recordset

       Set lRStFormes = gclsFormes.ConsultarFormsDeUmModuloComBotoes(vbytNumMod)

           vcboFormes.Clear
       Do _
           While (Not lRStFormes.EOF)
           vcboFormes.AddItem lRStFormes!NomFor & " - " & lRStFormes!NumFor
           lRStFormes.MoveNext
       Loop
       lRStFormes.Close
End Sub
Public Sub rgsCarregarFormesDeUmModuloComBotoesDeUmUsuario(ByVal vbytNumMod As Byte, ByVal vintNumUsu As Integer, _
                                                                                     ByRef vcboFormes As ComboBox)
       Dim lRStFormes As Recordset

       Set lRStFormes = gclsFormes.ConsultarFormsDeUmModuloComBotoesDeUmUsuario(vbytNumMod, vintNumUsu)

           vcboFormes.Clear
       Do _
           While (Not lRStFormes.EOF)
           vcboFormes.AddItem lRStFormes!NomFor & " - " & lRStFormes!Numero
           lRStFormes.MoveNext
       Loop
       lRStFormes.Close
End Sub
Public Sub rgsCarregarFundos(ByRef vcboFundos As ComboBox)
       Dim lRStFundos As Recordset

       Set lRStFundos = gclsFundos.ConsultarTodos

           vcboFundos.Clear
       Do _
           While (Not lRStFundos.EOF)
           vcboFundos.AddItem lRStFundos!Codigo & " - " & lRStFundos!Numero
           lRStFundos.MoveNext
       Loop
       lRStFundos.Close
End Sub
Public Sub rgsCarregarModulos(ByRef vcboModulo As ComboBox)
       Dim lRStModulo As Recordset

       Set lRStModulo = gclsModulo.ConsultarTodos

           vcboModulo.Clear
       Do _
           While (Not lRStModulo.EOF)
           vcboModulo.AddItem lRStModulo!Numero & " - " & lRStModulo!Descri
           lRStModulo.MoveNext
       Loop
       lRStModulo.Close
End Sub
Public Sub rgsCarregarModulosDeUmUsuario(ByVal vintNumUsu As Integer, ByRef vcboModulo As ComboBox)
       Dim lRStAceMod As Recordset

       Set lRStAceMod = gclsAceMod.ConsultarModulosDeUmUsuario(vintNumUsu)

           vcboModulo.Clear
       Do _
           While (Not lRStAceMod.EOF)
           vcboModulo.AddItem lRStAceMod!Numero & " - " & lRStAceMod!Descri
           lRStAceMod.MoveNext
       Loop
       lRStAceMod.Close
End Sub
Public Sub rgsCarregarUsuarios(ByRef vcboUsuari As ComboBox, ByVal vbooItem00 As Boolean)
       Dim lRStUsuari As Recordset

       Set lRStUsuari = gclsUsuari.ConsultarTodosPorNumero

           vcboUsuari.Clear

       If (vbooItem00) Then _
           vcboUsuari.AddItem "0 - Todos"
       Do _
           While (Not lRStUsuari.EOF)
           vcboUsuari.AddItem lRStUsuari!Numero & " - " & lRStUsuari!NomUsu
           lRStUsuari.MoveNext
       Loop
       lRStUsuari.Close
End Sub
Public Sub rgsCentralizarForm(ByRef vforFormes As Form)
       vforFormes.Left = (Screen.Width - vforFormes.Width) / 2
       vforFormes.Top = (formMDIAce.Height - vforFormes.Height - 1715) / 2
End Sub
Public Sub rgsCentralizarFormIndependente(ByRef vforFormes As Form)
       vforFormes.Left = (Screen.Width - vforFormes.Width) / 2
       vforFormes.Top = (Screen.Height - vforFormes.Height) / 2 + 200
End Sub
Public Sub rgsConsultarAdministradora()
       Dim lRStInsAdm As Recordset

       Set lRStInsAdm = gclsInsAdm.Consultar

       If (lRStInsAdm.EOF) Then
           gstrNomApl = "Acessos - Fundos"
       Else
           gstrNomApl = "Acessos - Fundos - " & lRStInsAdm!RazSoc
           gintCodAdm = lRStInsAdm!Codigo
       End If
       lRStInsAdm.Close
End Sub
Public Sub rgsLerUsuarioDoDataBase()
       Dim lintNumArq As Integer

       Dim lintPosIni As Integer

       Dim lstrArqLog As String

       Dim lstrLinArq As String

           lstrArqLog = "C:\Fundos\LogDBFun.Fun"

       If (Not ((rgfArquivoExiste(lstrArqLog)))) Then
           rgfMsgBox "Arquivo " & lstrArqLog & " não encontrado", MsgErr
           End
       End If

           gstrServBD = ""
           gstrNomeBD = ""
           gstrUsuaBD = ""
           gstrSenhBD = ""
           gstrPthExe = ""
           gstrPthAtu = ""
           gstrPthAju = ""
           gstrPthAAj = ""
           lintNumArq = FreeFile

           Open _
           lstrArqLog For Input As #lintNumArq
       Do _
           While (Not EOF(lintNumArq))
           Line Input #lintNumArq, lstrLinArq

           lintPosIni = InStr(1, lstrLinArq, ":", 1)

       If (lintPosIni = 0) Then lintPosIni = 2

           Select Case Mid(lstrLinArq, 1, lintPosIni - 1)
                  Case "ServBD", "SV"
                       gstrServBD = Trim(Mid(lstrLinArq, lintPosIni + 2))
                  Case "NomeBD", "BD"
                       gstrNomeBD = Trim(Mid(lstrLinArq, lintPosIni + 2))
                  Case "UsuaBD", "US"
                       gstrUsuaBD = Trim(Mid(lstrLinArq, lintPosIni + 2))
                  Case "SenhBD", "PW"
                       gstrSenhBD = Trim(Mid(lstrLinArq, lintPosIni + 2))

                  Case "VerAtu", "LE"
                       gstrPthExe = Trim(Mid(lstrLinArq, lintPosIni + 2))
                  Case "VerNov", "VE"
                       gstrPthAtu = Trim(Mid(lstrLinArq, lintPosIni + 2))

                  Case "LH"
                       gstrPthAju = Trim(Mid(lstrLinArq, lintPosIni + 2))
                  Case "VA"
                       gstrPthAAj = Trim(Mid(lstrLinArq, lintPosIni + 2))
           End Select
       Loop
       Close #lintNumArq

       gstrPthAtu = gstrPthAtu & "Acessos.exe"
       gstrPthAju = gstrPthAju & "Acessos.chm"
End Sub
Public Sub rgsPesquisarComboAll(ByRef vcboComboB As ComboBox, ByVal vstrItmCbo As String)
       vcboComboB.ListIndex = -1

       For gintNumItm = 0 To vcboComboB.ListCount - 1
       If (vstrItmCbo = vcboComboB.List(gintNumItm)) Then
           vcboComboB.ListIndex = gintNumItm
           Exit For
       End If
       Next
End Sub
Public Sub rgsPesquisarComboFim(ByRef vcboComboB As ComboBox, ByVal vstrItmCbo As String)
       vcboComboB.ListIndex = -1

       For gintNumItm = 0 To vcboComboB.ListCount - 1
       If (vstrItmCbo = Mid(vcboComboB.List(gintNumItm), InStrRev(vcboComboB.List(gintNumItm), "-") + 2, Len(vcboComboB.List(gintNumItm)) - InStrRev(vcboComboB.List(gintNumItm), "-") + 1)) Then
           vcboComboB.ListIndex = gintNumItm
           Exit For
       End If
       Next
End Sub
Public Sub rgsPesquisarComboIni(ByRef vcboComboB As ComboBox, ByVal vstrItmCbo As String)
       vcboComboB.ListIndex = -1

       For gintNumItm = 0 To vcboComboB.ListCount - 1
       If (vstrItmCbo = Mid(vcboComboB.List(gintNumItm), 1, InStr(vcboComboB.List(gintNumItm), " ") - 1)) Then
           vcboComboB.ListIndex = gintNumItm
           Exit For
       End If
       Next
End Sub
Public Sub rgsPosicionarAjuda(ByRef vforFormes As Form, ByRef vintForAtu As Integer, ByRef vbooForLog As Boolean)
       Dim lRStFormes As Recordset

       Set lRStFormes = gclsFormes.ConsultarParaAjuda(1, vforFormes.Name)

       If (lRStFormes.EOF) Then
           vintForAtu = 0
           vbooForLog = False
           vforFormes.HelpContextID = 1
       Else
           vintForAtu = lRStFormes!Numero
'          vbooForLog = IIf( _
'                       lRStFormes!TemLog, 1, 0)
           vforFormes.HelpContextID = lRStFormes!NumAju
       End If
       lRStFormes.Close
End Sub
Public Sub rgsTratarFuncoes(ByVal vintKeyCod As Integer, ByRef vforFormes As Form)
       Dim lstrCmbNom As String

       If (vintKeyCod < 112 Or vintKeyCod > 123) And vintKeyCod <> 27 Then Exit Sub

           lstrCmbNom = "fcmb" & IIf(vintKeyCod = vbKeyEscape, "Escape", "F" & Format(vintKeyCod - 111, "00"))

       For gintNumItm = 0 To vforFormes.Controls.Count - 1
       If (InStr(1, vforFormes.Controls(gintNumItm).Name, lstrCmbNom, vbTextCompare)) Then
       If (vforFormes.Controls(gintNumItm).Enabled) Then
           vforFormes.Controls(gintNumItm).Value = True
       End If
           Exit For
       End If
       Next
End Sub
