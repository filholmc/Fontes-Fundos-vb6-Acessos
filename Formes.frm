VERSION 5.00
Begin VB.Form formFormes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forms"
   ClientHeight    =   2820
   ClientLeft      =   4185
   ClientTop       =   1755
   ClientWidth     =   5910
   Icon            =   "Formes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "Formes.frx":030A
   ScaleHeight     =   2820
   ScaleWidth      =   5910
   Begin VB.CheckBox fchkTemLog 
      Caption         =   "Tem Log"
      Height          =   195
      Left            =   2115
      TabIndex        =   5
      Top             =   1890
      Width           =   945
   End
   Begin VB.CheckBox fchkSemAce 
      Caption         =   "Acesso Livre"
      Height          =   195
      Left            =   570
      TabIndex        =   4
      Top             =   1890
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   0
      TabIndex        =   21
      Top             =   2250
      Width           =   5895
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   510
      TabIndex        =   20
      Top             =   450
      Width           =   5385
   End
   Begin VB.TextBox ftxtDescri 
      Height          =   315
      Left            =   570
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1410
      Width           =   4785
   End
   Begin VB.TextBox ftxtNomFor 
      Height          =   315
      Left            =   1500
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
      ItemData        =   "Formes.frx":0BD4
      Left            =   3570
      List            =   "Formes.frx":0BD6
      TabIndex        =   2
      Top             =   810
      Width           =   1800
   End
   Begin VB.TextBox ftxtNumAju 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4920
      MaxLength       =   2
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1830
      Width           =   435
   End
   Begin VB.CommandButton fcmbF09LCA 
      Caption         =   "F9"
      Height          =   255
      Left            =   4470
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF11Tab 
      Caption         =   "F11"
      Height          =   255
      Left            =   4920
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF12Hom 
      Caption         =   "F12"
      Height          =   255
      Left            =   5370
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF06Loc 
      Caption         =   "Localizar (F6)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3540
      TabIndex        =   10
      Top             =   2400
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF05Exc 
      Caption         =   "Excluir (F5)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2400
      TabIndex        =   9
      Top             =   2400
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF03Alt 
      Caption         =   "Alterar (F3)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1260
      TabIndex        =   8
      Top             =   2400
      Width           =   1120
   End
   Begin VB.CommandButton fcmbEscape 
      Caption         =   "Fechar (Esc)"
      Height          =   315
      Left            =   4680
      TabIndex        =   11
      Top             =   2400
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF02Inc 
      Caption         =   "Incluir (F2)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   2400
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
      TabIndex        =   19
      Top             =   600
      Width           =   660
   End
   Begin VB.Label flblDescri 
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
      Height          =   195
      Left            =   570
      LinkTimeout     =   0
      TabIndex        =   18
      Top             =   1215
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nº da Tela de Ajuda"
      Height          =   195
      Left            =   3390
      TabIndex        =   17
      Top             =   1890
      Width           =   1440
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      Height          =   195
      Left            =   1500
      TabIndex        =   16
      Top             =   600
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Módulo"
      Height          =   195
      Left            =   3570
      TabIndex        =   15
      Top             =   600
      Width           =   525
   End
End
Attribute VB_Name = "formFormes"
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

Private fbytNumAju As Byte

Private fbytNumMod As Byte

Private fclsContro As clssContro

Private fclsBotoes As clssBotoes

Private fclsAceFor As clssAceFor

Private fintForAtu As Integer

Private fintNumero As Integer, fintNumFor As Integer
Private Function rlfValidacaoDeCamposOk() As Boolean
        rlfValidacaoDeCamposOk = False

        If (fcboModulo.ListIndex = -1) Then
            rgfMsgBox "Escolha uma opção do campo 'Módulo'", MsgErr
            fcboModulo.SetFocus
            Exit Function
        End If

        If (Trim(ftxtDescri) = "") Then
            rgfMsgBox "Preencha o campo 'Descrição'", MsgErr, Me.HelpContextID
            ftxtDescri.SetFocus
            Exit Function
        End If

        If (Trim(ftxtNomFor) = "") Then
            rgfMsgBox "Preencha o campo 'Nome'", MsgErr, Me.HelpContextID
            ftxtNomFor.SetFocus
            Exit Function
        End If

            rlsConsultarNome

        If (fbooJahCad) And _
           (fintNumero) <> fintNumFor Then
            rgfMsgBox "Nome já utilizado neste Módulo", MsgErr
            ftxtNomFor.SetFocus
            Exit Function
        End If

        If (ftxtNumAju = "") Then
            ftxtNumAju = 0
        End If

            fbytNumAju = CByte(ftxtNumAju)

            rlsConsultarAjuda

        If (fbooJahCad And fintNumero <> fintNumFor) Then
            rgfMsgBox "Número da Tela de Ajuda já utilizado neste Módulo", MsgErr
            ftxtNumAju.SetFocus
            Exit Function
        End If

        rlfValidacaoDeCamposOk = True
End Function
Private Sub fcboModulo_Click()
        If (fcboModulo.ListIndex = -1) Then Exit Sub

        fbytNumMod = CByte(Mid(fcboModulo, 1, InStr(fcboModulo, " ") - 1))
End Sub
Private Sub fcmbEscape_Click()
        Unload Me
End Sub
Private Sub fcmbF02Inc_Click()
        If (Not rlfValidacaoDeCamposOk) Then Exit Sub

        rlsConsultarControle

        gDBCFundos.BeginTrans
                   fclsContro.Alterar "UltFor", fintNumero

               If (gDBCFundos.Errors.Count > 0) Then GoTo Erro_DB

                   gclsFormes.Incluir fbytNumMod, fintNumero, ftxtNomFor, ftxtDescri, fchkSemAce, fchkTemLog, fbytNumAju

               If (gDBCFundos.Errors.Count > 0) Then GoTo Erro_DB

                   gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                               "Incluiu", 0, Format(fintNumero, "0"), _
                                                                             ftxtNomFor & "; " & _
                                                                             Format(fbytNumMod, "0") & "; " & _
                                                                             ftxtDescri & "; " & fchkSemAce & "; " & _
                                                                             fchkTemLog & "; " & _
                                                                             Format(fbytNumAju, "0")
               If (gDBCFundos.Errors.Count > 0) Then GoTo Erro_DB
        gDBCFundos.CommitTrans

        rgfMsgBox "Form Incluído", MsgInf

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

            gclsFormes.Alterar fbytNumMod, fintNumero, ftxtNomFor, ftxtDescri, fchkSemAce, fchkTemLog, fbytNumAju

            gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                        "Alterou", 0, Format(fintNumero, "0"), _
                                                                      ftxtNomFor & "; " & _
                                                                      Format(fbytNumMod, "0") & "; " & _
                                                                      ftxtDescri & "; " & fchkSemAce & "; " & _
                                                                      fchkTemLog & "; " & _
                                                                      Format(fbytNumAju, "0")
        If (gDBCFundos.Errors.Count > 0) Then
            rgsTratarErro Err, gDBCFundos.Errors, Me
        Else
            rgfMsgBox "Dados do Form alterados", MsgInf
            ftxtNumero.SetFocus
        End If
End Sub
Private Sub fcmbF05Exc_Click()
        If (rgfMsgBox("Confirma Exclusão?", MsgNao) = vbYes) Then
            gclsFormes.Excluir fintNumero
            gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                        "Excluiu", 0, Format(fintNumero, "0"), " "
        If (gDBCFundos.Errors.Count > 0) Then
            rgsTratarErro Err, gDBCFundos.Errors, Me
        Else
            rgfMsgBox "Form Excluído", MsgInf
            fcmbF03Alt.Enabled = False
            fcmbF05Exc.Enabled = False
        End If
        End If
        ftxtNumero.SetFocus
End Sub
Private Sub fcmbF06Loc_Click()
        formConFor.SetFocus
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
        If (Not gbooConFor) Then Exit Sub

        gbooConFor = False
        ftxtNumero = gintNumFor
        ftxtNumero_LostFocus
        ftxtNumero.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        gintTempoP = 0
        rgsTratarFuncoes KeyCode, Me
End Sub
Private Sub Form_Load()
        formMDIAce.ftbrModAce.Buttons("Formes").Value = tbrPressed

        rgsCentralizarForm Me
        rgsPosicionarAjuda Me, fintForAtu, fbooForLog

        Set fclsContro = New clssContro
        Set fclsBotoes = New clssBotoes
        Set fclsAceFor = New clssAceFor

        rgsCarregarModulos fcboModulo

        rlsHabilitarBotoes

        fcmbF06Loc.Enabled = IIf(formMDIAce.menuConFor.Enabled, True, False)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        gintTempoP = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
        Set fclsContro = Nothing
        Set fclsBotoes = Nothing
        Set fclsAceFor = Nothing
        formMDIAce.fsbrModAce.Panels(4).Picture = LoadPicture()
        formMDIAce.ftbrModAce.Buttons("Formes").Value = tbrUnpressed
End Sub
Private Sub ftxtNumAju_GotFocus()
        ftxtNumAju.SelStart = Len(ftxtNumAju)
End Sub
Private Sub ftxtNumAju_KeyPress(KeyAscii As Integer)
        If (Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8) Then KeyAscii = 0
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
        Dim lRStFormes As Recordset

        Set lRStFormes = gclsFormes.Consultar(fintNumero)

        If (lRStFormes.EOF) Then
            rlsLimparCampos
        Else
            rgsPesquisarComboIni fcboModulo, lRStFormes!NumMod

            ftxtNomFor = lRStFormes!NomFor
            fbytNumMod = lRStFormes!NumMod
            ftxtDescri = lRStFormes!Descri
            fchkSemAce = IIf( _
                         lRStFormes!SemAce, 1, 0)
            fchkTemLog = IIf( _
                         lRStFormes!TemLog, 1, 0)
            ftxtNumAju = lRStFormes!NumAju

            fchkTemLog.Enabled = Not _
                                 lRStFormes!SemLog
            fcmbF02Inc.Enabled = False
            fcmbF03Alt.Enabled = fbooF03Alt
            fcmbF05Exc.Enabled = IIf(fclsBotoes.PertenceForm(fintNumero) Or fclsAceFor.FormAcessado(fintNumero), False, fbooF05Exc)
        End If
        lRStFormes.Close
End Sub
Private Sub rlsConsultarAjuda()
        Dim lRStFormes As Recordset

        Set lRStFormes = gclsFormes.ConsultarAjuda(fbytNumMod, fbytNumAju)

        If (lRStFormes.EOF) Then
            fbooJahCad = False
        Else
            fbooJahCad = True
            fintNumFor = lRStFormes!Numero
        End If
        lRStFormes.Close
End Sub
Private Sub rlsConsultarControle()
        Dim lRStContro As Recordset

        Set lRStContro = fclsContro.Consultar

            fintNumero = lRStContro!UltFor + 1

        lRStContro.Close
End Sub
Private Sub rlsConsultarNome()
        Dim lRStFormes As Recordset

        Set lRStFormes = gclsFormes.ConsultarNome(fbytNumMod, ftxtNomFor)

        If (lRStFormes.EOF) Then
            fbooJahCad = False
        Else
            fbooJahCad = True
            fintNumFor = lRStFormes!Numero
        End If
        lRStFormes.Close
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
        ftxtNomFor = ""
        fcboModulo = ""
        ftxtDescri = ""
        fchkSemAce = 0
        fchkTemLog = 0
        ftxtNumAju = ""
        fchkSemAce.Enabled = IIf(gintCodAdm = 7, True, False)
        fcmbF02Inc.Enabled = fbooF02Inc
        fcmbF03Alt.Enabled = False
        fcmbF05Exc.Enabled = False
End Sub
