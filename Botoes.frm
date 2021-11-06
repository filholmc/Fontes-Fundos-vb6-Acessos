VERSION 5.00
Begin VB.Form formBotoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Botões"
   ClientHeight    =   3060
   ClientLeft      =   4125
   ClientTop       =   1755
   ClientWidth     =   5910
   Icon            =   "Botoes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "Botoes.frx":030A
   ScaleHeight     =   3060
   ScaleWidth      =   5910
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   0
      TabIndex        =   19
      Top             =   2490
      Width           =   5895
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   510
      TabIndex        =   18
      Top             =   450
      Width           =   5385
   End
   Begin VB.TextBox ftxtDescri 
      Height          =   315
      Left            =   570
      MaxLength       =   30
      TabIndex        =   4
      Top             =   2070
      Width           =   4755
   End
   Begin VB.TextBox ftxtNomBot 
      Height          =   315
      Left            =   3540
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   1
      Top             =   810
      Width           =   1800
   End
   Begin VB.TextBox ftxtNumero 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   570
      MaxLength       =   5
      MultiLine       =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Na Inclusão, o Número é gerado automaticamente"
      Top             =   810
      Width           =   645
   End
   Begin VB.ComboBox fcboModulo 
      Height          =   315
      ItemData        =   "Botoes.frx":0BD4
      Left            =   570
      List            =   "Botoes.frx":0BD6
      TabIndex        =   2
      Top             =   1440
      Width           =   1800
   End
   Begin VB.ComboBox fcboFormes 
      Height          =   315
      ItemData        =   "Botoes.frx":0BD8
      Left            =   3540
      List            =   "Botoes.frx":0BDA
      TabIndex        =   3
      Top             =   1440
      Width           =   1800
   End
   Begin VB.CommandButton fcmbF09LCA 
      Caption         =   "F9"
      Height          =   255
      Left            =   4470
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF11Tab 
      Caption         =   "F11"
      Height          =   255
      Left            =   4920
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF12Hom 
      Caption         =   "F12"
      Height          =   255
      Left            =   5370
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF06Loc 
      Caption         =   "Localizar (F6)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3540
      TabIndex        =   8
      Top             =   2640
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF05Exc 
      Caption         =   "Excluir (F5)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Top             =   2640
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF03Alt 
      Caption         =   "Alterar (F3)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1260
      TabIndex        =   6
      Top             =   2640
      Width           =   1120
   End
   Begin VB.CommandButton fcmbEscape 
      Caption         =   "Fechar (Esc)"
      Height          =   315
      Left            =   4680
      TabIndex        =   9
      Top             =   2640
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF02Inc 
      Caption         =   "Incluir (F2)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1120
   End
   Begin VB.Label flblNumero 
      AutoSize        =   -1  'True
      Caption         =   "Número"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   570
      LinkTimeout     =   0
      TabIndex        =   17
      Top             =   600
      Width           =   660
   End
   Begin VB.Label flblDescri 
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
      Height          =   195
      Left            =   570
      LinkTimeout     =   0
      TabIndex        =   16
      Top             =   1860
      Width           =   750
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      Height          =   195
      Left            =   3540
      TabIndex        =   15
      Top             =   600
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Módulo"
      Height          =   195
      Left            =   570
      TabIndex        =   14
      Top             =   1230
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Form"
      Height          =   195
      Left            =   3540
      TabIndex        =   13
      Top             =   1230
      Width           =   345
   End
End
Attribute VB_Name = "formBotoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private fbooF02Inc As Boolean

Private fbooF03Alt As Boolean

Private fbooF05Exc As Boolean

Private fbooForLog As Boolean

Private fbooJahCad As Boolean

Private fbytNumMod As Byte

Private fclsContro As clssContro

Private fclsBotoes As clssBotoes

Private fintForAtu As Integer

Private fintNumBot As Integer, fintNumero As Integer, fintNumFor As Integer
Private Function rlfValidacaoDeCamposOk() As Boolean
        rlfValidacaoDeCamposOk = False

        If (Trim(ftxtNomBot) = "") Then
            rgfMsgBox "Preencha o campo 'Nome'", MsgErr, Me.HelpContextID
            ftxtNomBot.SetFocus
            Exit Function
        End If

            rlsConsultarNome

        If (fbooJahCad) And _
           (fintNumero <> fintNumBot) Then
            rgfMsgBox "Nome já utilizado neste Form", MsgErr
            ftxtNomBot.SetFocus
            Exit Function
        End If

        If (fcboModulo.ListIndex = -1) Then
            rgfMsgBox "Escolha uma opção do campo 'Módulo'", MsgErr
            fcboModulo.SetFocus
            Exit Function
        End If

        If (fcboFormes.ListIndex = -1) Then
            rgfMsgBox "Escolha uma opção do campo 'Form'", MsgErr
            fcboFormes.SetFocus
            Exit Function
        End If

        If (Trim(ftxtDescri)) = "" Then
            rgfMsgBox "Preencha o campo 'Descrição'", MsgErr, Me.HelpContextID
            ftxtDescri.SetFocus
            Exit Function
        End If

        rlfValidacaoDeCamposOk = True
End Function
Private Sub fcboFormes_Click()
        If (fcboFormes.ListIndex = -1) Then Exit Sub

        fintNumFor = CInt(Mid(fcboFormes, 14, Len(fcboFormes) - 13))
End Sub
Private Sub fcboModulo_Click()
        If (fcboModulo.ListIndex = -1) Then Exit Sub

        fbytNumMod = CByte(Mid(fcboModulo, 1, InStr(fcboModulo, " ") - 1))

        rgsCarregarFormesDeUmModuloAcessaveis fbytNumMod, fcboFormes, False
End Sub
Private Sub fcmbEscape_Click()
        Unload Me
End Sub
Private Sub fcmbF02Inc_Click()
        If (Not rlfValidacaoDeCamposOk) Then Exit Sub

        rlsConsultarControle

        gDBCFundos.BeginTrans
                   fclsContro.Alterar "UltBot", fintNumero

               If (gDBCFundos.Errors.Count > 0) Then GoTo Erro_DB

                   fclsBotoes.Incluir fbytNumMod, fintNumFor, fintNumero, ftxtNomBot, ftxtDescri

               If (gDBCFundos.Errors.Count > 0) Then GoTo Erro_DB

                   gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                               "Incluiu", 0, Format(fintNumero, "0"), _
                                                                             ftxtNomBot & "; " & _
                                                                             Format(fbytNumMod, "0") & "; " & _
                                                                             Format(fintNumFor, "0") & "; " & ftxtDescri
               If (gDBCFundos.Errors.Count > 0) Then GoTo Erro_DB
        gDBCFundos.CommitTrans

        rgfMsgBox "Botão Incluído", MsgInf

        fcmbF02Inc.Enabled = False
        fcmbF03Alt.Enabled = fbooF03Alt
        fcmbF05Exc.Enabled = fbooF05Exc
        ftxtNumero = fintNumero
        ftxtNumero.SetFocus
        Exit Sub

Erro_DB:
        gDBCFundos.RollbackTrans

        rgsTratarErro Err, gDBCFundos.Errors, Me
End Sub
Private Sub fcmbF03Alt_Click()
        If (Not rlfValidacaoDeCamposOk) Then Exit Sub

            fclsBotoes.Alterar fbytNumMod, fintNumFor, fintNumero, ftxtNomBot, ftxtDescri

            gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                        "Alterou", 0, Format(fintNumero, "0"), _
                                                                      ftxtNomBot & "; " & _
                                                                      Format(fbytNumMod, "0") & "; " & _
                                                                      Format(fintNumFor, "0") & "; " & ftxtDescri
        If (gDBCFundos.Errors.Count > 0) Then
            rgsTratarErro Err, gDBCFundos.Errors, Me
        Else
            rgfMsgBox "Dados do Botão alterados", MsgInf
        End If

        ftxtNumero.SetFocus
End Sub
Private Sub fcmbF05Exc_Click()
        If (rgfMsgBox("Confirma Exclusão?", MsgNao) = vbYes) Then
            fclsBotoes.Excluir fintNumero
            gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                        "Excluiu", 0, Format(fintNumero, "0"), " "
        If (gDBCFundos.Errors.Count > 0) Then
            rgsTratarErro Err, gDBCFundos.Errors, Me
        Else
            rgfMsgBox "Botão Excluído", MsgInf
            fcmbF03Alt.Enabled = False
            fcmbF05Exc.Enabled = False
        End If
        End If
        ftxtNumero.SetFocus
End Sub
Private Sub fcmbF06Loc_Click()
        formConBot.SetFocus
End Sub
Private Sub fcmbF09LCA_Click()
        If (Not TypeOf ActiveControl Is CommandButton) Then ActiveControl.Text = ""
End Sub
Private Sub fcmbF11Tab_Click()
        SendKeys "+{TAB}"
End Sub
Private Sub fcmbF12Hom_Click()
        ftxtNumero.SetFocus
End Sub
Private Sub Form_Activate()
        gintTempoP = 0
        formMDIAce.fsbrModAce.Panels(4).Picture = IIf(gbooUsuLog Or fbooForLog, _
                                                      formMDIAce.fimlStaBar.ListImages(1).Picture, LoadPicture())
        If (Not gbooConBot) Then Exit Sub

        gbooConBot = False
        ftxtNumero = gintNumBot
        ftxtNumero_LostFocus
        ftxtNumero.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        gintTempoP = 0
        rgsTratarFuncoes KeyCode, Me
End Sub
Private Sub Form_Load()
        formMDIAce.ftbrModAce.Buttons("Botoes").Value = tbrPressed

        rgsCentralizarForm Me
        rgsPosicionarAjuda Me, fintForAtu, fbooForLog

        Set fclsContro = New clssContro
        Set fclsBotoes = New clssBotoes

        rgsCarregarModulos fcboModulo

        rlsHabilitarBotoes

        fcmbF06Loc.Enabled = IIf(formMDIAce.menuConBot.Enabled, True, False)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        gintTempoP = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
        Set fclsContro = Nothing
        Set fclsBotoes = Nothing
        formMDIAce.fsbrModAce.Panels(4).Picture = LoadPicture()
        formMDIAce.ftbrModAce.Buttons("Botoes").Value = tbrUnpressed
End Sub
Private Sub ftxtNumero_GotFocus()
        rlsDesabilitarBotoes
        ftxtNumero.SelStart = Len(ftxtNumero)
End Sub
Private Sub ftxtNumero_KeyPress(KeyAscii As Integer)
        If (Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8) Then KeyAscii = 0
End Sub
Private Sub ftxtNumero_KeyUp(KeyCode As Integer, Shift As Integer)
        If (KeyCode = 8) Or (KeyCode >= 48 And KeyCode <= 57) Or (KeyCode >= 96 And KeyCode <= 105) Then
            ftxtNumero = Format(ftxtNumero, "#,##0")
            ftxtNumero.SelStart = Len(ftxtNumero)
        End If
End Sub
Private Sub ftxtNumero_LostFocus()
        rlsFormatarChaves
        rlsConsultar
End Sub
Private Sub rlsConsultar()
        Dim lRStBotoes As Recordset

        Set lRStBotoes = fclsBotoes.Consultar(fintNumero)

        If (lRStBotoes.EOF) Then
            rlsLimparCampos
        Else
            rgsCarregarFormesDeUmModuloAcessaveis lRStBotoes!NumMod, fcboFormes, False

            rgsPesquisarComboIni fcboModulo, lRStBotoes!NumMod
            rgsPesquisarComboFim fcboFormes, lRStBotoes!NumFor

            fbytNumMod = lRStBotoes!NumMod
            ftxtNomBot = lRStBotoes!NomBot
            ftxtDescri = lRStBotoes!Descri

            fcmbF02Inc.Enabled = False
            fcmbF03Alt.Enabled = fbooF03Alt
            fcmbF05Exc.Enabled = IIf(gclsAceBot.BotaoAcessado(fintNumero), False, fbooF05Exc)
        End If
        lRStBotoes.Close
End Sub
Private Sub rlsConsultarControle()
        Dim lRStContro As Recordset

        Set lRStContro = fclsContro.Consultar

            fintNumero = lRStContro!UltBot + 1

        lRStContro.Close
End Sub
Private Sub rlsConsultarNome()
        Dim lRStBotoes As Recordset

        Set lRStBotoes = fclsBotoes.ConsultarNome(fbytNumMod, fintNumFor, ftxtNomBot)

        If (lRStBotoes.EOF) Then
            fbooJahCad = False
        Else
            fbooJahCad = True
            fintNumBot = lRStBotoes!Numero
        End If
        lRStBotoes.Close
End Sub
Private Sub rlsDesabilitarBotoes()
        fcmbF03Alt.Enabled = False
        fcmbF05Exc.Enabled = False
End Sub
Private Sub rlsFormatarChaves()
        If (ftxtNumero = "") Then
            ftxtNumero = 0
        End If

        fintNumero = CInt(rgfSemEdicao(ftxtNumero))
End Sub
Private Sub rlsHabilitarBotoes()
        Dim lRStAceBot As Recordset

        Set lRStAceBot = gclsAceBot.ConsultarBotoesDeUmUsuarioPorModuloAndForm(gintUsuLog, 1, fintForAtu)

        Do _
            While (Not ((lRStAceBot.EOF)))
            Select Case (lRStAceBot!NomBot)
                   Case "fcmbF02Inc"
                         fbooF02Inc = True
                   Case "fcmbF03Alt"
                         fbooF03Alt = True
                   Case "fcmbF05Exc"
                         fbooF05Exc = True
            End Select
            lRStAceBot.MoveNext
        Loop
        lRStAceBot.Close
End Sub
Private Sub rlsLimparCampos()
        ftxtNomBot = ""
        fcboModulo = ""
        fcboFormes = ""
        ftxtDescri = ""
        fcmbF02Inc.Enabled = fbooF02Inc
        fcmbF03Alt.Enabled = False
        fcmbF05Exc.Enabled = False
End Sub
