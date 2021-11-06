VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form formUsuari 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuários"
   ClientHeight    =   4260
   ClientLeft      =   4140
   ClientTop       =   1755
   ClientWidth     =   5910
   Icon            =   "Usuari.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "Usuari.frx":030A
   ScaleHeight     =   4260
   ScaleWidth      =   5910
   Begin VB.CheckBox fchkTemLog 
      Caption         =   "Tem Log"
      Height          =   195
      Left            =   2145
      TabIndex        =   1
      Top             =   810
      Width           =   945
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   0
      TabIndex        =   24
      Top             =   3690
      Width           =   5895
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   510
      TabIndex        =   23
      Top             =   390
      Width           =   5385
   End
   Begin VB.TextBox ftxtNomUsu 
      Height          =   315
      Left            =   570
      MaxLength       =   40
      TabIndex        =   3
      Top             =   1380
      Width           =   4755
   End
   Begin VB.TextBox ftxtE_Mail 
      Height          =   315
      Left            =   570
      LinkTimeout     =   0
      MaxLength       =   40
      TabIndex        =   7
      Top             =   3270
      Width           =   4755
   End
   Begin VB.ComboBox fcboCodAge 
      Height          =   315
      Left            =   570
      TabIndex        =   4
      Top             =   2010
      Width           =   4755
   End
   Begin VB.TextBox ftxtNumero 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   570
      MaxLength       =   5
      MultiLine       =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Na Inclusão, o Número é gerado automaticamente"
      Top             =   750
      Width           =   645
   End
   Begin VB.ComboBox fcboFuncao 
      Height          =   315
      Left            =   4020
      TabIndex        =   6
      Top             =   2640
      Width           =   1305
   End
   Begin VB.ComboBox fcboStatus 
      Height          =   315
      ItemData        =   "Usuari.frx":0BD4
      Left            =   4020
      List            =   "Usuari.frx":0BDE
      TabIndex        =   2
      Top             =   750
      Width           =   1305
   End
   Begin VB.CommandButton fcmbF09LCA 
      Caption         =   "F9"
      Height          =   255
      Left            =   4470
      TabIndex        =   13
      Top             =   90
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF11Tab 
      Caption         =   "F11"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   90
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF12Hom 
      Caption         =   "F12"
      Height          =   255
      Left            =   5370
      TabIndex        =   15
      Top             =   90
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF06Loc 
      Caption         =   "Localizar (F6)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3540
      TabIndex        =   11
      Top             =   3840
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF05Exc 
      Caption         =   "Excluir (F5)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2400
      TabIndex        =   10
      Top             =   3840
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF03Alt 
      Caption         =   "Alterar (F3)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1260
      TabIndex        =   9
      Top             =   3840
      Width           =   1120
   End
   Begin VB.CommandButton fcmbEscape 
      Caption         =   "Fechar (Esc)"
      Height          =   315
      Left            =   4680
      TabIndex        =   12
      Top             =   3840
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF02Inc 
      Caption         =   "Incluir (F2)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   1120
   End
   Begin MSMask.MaskEdBox fmskDatVal 
      Height          =   315
      Left            =   570
      TabIndex        =   5
      ToolTipText     =   "Data até a qual o Usuário terá Acesso ao Sistema"
      Top             =   2640
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   4
      Mask            =   "##/##/####"
      PromptChar      =   " "
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
      TabIndex        =   22
      Top             =   540
      Width           =   660
   End
   Begin VB.Label flblNomUsu 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      Height          =   195
      Left            =   570
      LinkTimeout     =   0
      TabIndex        =   21
      Top             =   1170
      Width           =   420
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "e-mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   570
      LinkTimeout     =   0
      TabIndex        =   20
      Top             =   3060
      Width           =   495
   End
   Begin VB.Label flblNomAge 
      AutoSize        =   -1  'True
      Caption         =   "Agência"
      Height          =   195
      Left            =   570
      TabIndex        =   19
      Top             =   1800
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Função"
      Height          =   195
      Left            =   4020
      TabIndex        =   18
      Top             =   2430
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Acesso até"
      Height          =   195
      Left            =   570
      TabIndex        =   17
      Top             =   2430
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Left            =   4020
      TabIndex        =   16
      Top             =   540
      Width           =   450
   End
End
Attribute VB_Name = "formUsuari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private fbooF02Inc As Boolean

Private fbooF03Alt As Boolean

Private fbooF05Exc As Boolean

Private fbooForLog As Boolean

Private fbytIndice As Byte

Private fclsContro As clssContro

Private fclsAceFun As clssAceFun

Private fclsAceFor As clssAceFor

Private fclsAgeCad As clssAgeCad

Private fdatDatVal As Date

Private fintForAtu As Integer

Private fintCodAge As Integer, fintNumero As Integer
Private Function rlfValidacaoDeCamposOk() As Boolean
        rlfValidacaoDeCamposOk = False

        If (fcboStatus.ListIndex = -1) Then
            rgfMsgBox "Escolha uma opção do campo 'Status'", MsgErr
            fcboStatus.SetFocus
            Exit Function
        End If

        If (Trim(ftxtNomUsu) = "") Then
            rgfMsgBox "Preencha o campo 'Nome'", MsgErr, Me.HelpContextID
            ftxtNomUsu.SetFocus
            Exit Function
        End If

        If (fcboCodAge.ListIndex = -1) Then
            rgfMsgBox "Escolha uma opção do campo 'Agência'", MsgErr
            fcboCodAge.SetFocus
            Exit Function
        End If

        If (Not IsDate(Format(fmskDatVal, "00/00/0000"))) Then
            rgfMsgBox "Corrija o campo 'Acesso até'", MsgErr, Me.HelpContextID
            fmskDatVal.SetFocus
            Exit Function
        End If

            fdatDatVal = Format(fmskDatVal, "00/00/0000")

        If (fdatDatVal < gdatServBD) Then
            rgfMsgBox "Usuário deve ter pelo menos 1 Dia de Acesso ao Sistema", MsgErr
            fmskDatVal.SetFocus
            Exit Function
        End If

        If (fcboFuncao.ListIndex = -1) Then
            rgfMsgBox "Escolha uma opção do campo 'Função'", MsgErr
            fcboFuncao.SetFocus
            Exit Function
        End If

        rlfValidacaoDeCamposOk = True
End Function
Private Sub fcboCodAge_Click()
        If (fcboCodAge.ListIndex = -1) Then Exit Sub

            fintCodAge = CInt(Mid(fcboCodAge, 1, InStr(fcboCodAge, " ") - 1))

        If (fintCodAge = 0) Then
            rlsCarregarFuncoesDeDirecao
        Else
            rlsCarregarFuncoesDeAgencia
        End If
End Sub
Private Sub fcmbEscape_Click()
        Unload Me
End Sub
Private Sub fcmbF02Inc_Click()
        If (Not rlfValidacaoDeCamposOk) Then Exit Sub

        rlsConsultarControle

        gDBCFundos.BeginTrans
                   fclsContro.Alterar "UltUsu", fintNumero

               If (gDBCFundos.Errors.Count > 0) Then GoTo Erro_DB

                   gclsUsuari.Incluir fintNumero, rgfSenhaCp("*"), fchkTemLog, fcboStatus.ListIndex, _
                                                                              ftxtNomUsu, fintCodAge, _
                                                                              fdatDatVal, fcboFuncao, _
                                                                              ftxtE_Mail
               If (gDBCFundos.Errors.Count > 0) Then GoTo Erro_DB

                   gclsDiario.Incluir fbooForLog, gintUsuLog, _
                                                               gstrNomCmp, 1, fintForAtu, _
                                                                "Incluiu", 0, Format(fintNumero, "0"), _
                                                                              fchkTemLog & "; " & fcboStatus & "; " & _
                                                                              ftxtNomUsu & "; " & _
                                                                              Format(fintCodAge, "0") & "; " & _
                                                                              Format(fdatDatVal, "dd/mm/yyyy") & "; " & _
                                                                              fcboFuncao & "; " & ftxtE_Mail
               If (gDBCFundos.Errors.Count > 0) Then GoTo Erro_DB
        gDBCFundos.CommitTrans

        rgfMsgBox "Usuário Incluído", MsgInf

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

            gclsUsuari.Alterar fintNumero, fchkTemLog, fcboStatus.ListIndex, ftxtNomUsu, fintCodAge, fdatDatVal, _
                                                                                         fcboFuncao, ftxtE_Mail
            gclsDiario.Incluir fbooForLog, gintUsuLog, _
                                                              gstrNomCmp, 1, fintForAtu, _
                                                               "Alterou", 0, Format(fintNumero, "0"), _
                                                                             fchkTemLog & "; " & fcboStatus & "; " & _
                                                                             ftxtNomUsu & "; " & _
                                                                             Format(fintCodAge, "0") & "; " & _
                                                                             Format(fdatDatVal, "dd/mm/yyyy") & "; " & _
                                                                             fcboFuncao & "; " & ftxtE_Mail
        If (gDBCFundos.Errors.Count > 0) Then
            rgsTratarErro Err, gDBCFundos.Errors, Me
        Else
            rgfMsgBox "Dados do Usuário alterados", MsgInf
            ftxtNumero.SetFocus
        End If
End Sub
Private Sub fcmbF05Exc_Click()
        If (rgfMsgBox("Confirma Exclusão?", MsgNao) = vbYes) Then
            gclsUsuari.Excluir fintNumero
            gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                        "Excluiu", 0, Format(fintNumero, "0"), " "
        If (gDBCFundos.Errors.Count > 0) Then
            rgsTratarErro Err, gDBCFundos.Errors, Me
        Else
            rgfMsgBox "Usuário Excluído", MsgInf
            fcmbF03Alt.Enabled = False
            fcmbF05Exc.Enabled = False
        End If
        End If
        ftxtNumero.SetFocus
End Sub
Private Sub fcmbF06Loc_Click()
        formConUsu.SetFocus
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
        If (Not gbooConUsu) Then Exit Sub

        gbooConUsu = False
        ftxtNumero = gintNumUsu
        ftxtNumero_LostFocus
        ftxtNumero.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        gintTempoP = 0
        rgsTratarFuncoes KeyCode, Me
End Sub
Private Sub Form_Load()
        formMDIAce.ftbrModAce.Buttons("Usuari").Value = tbrPressed

        rgsCentralizarForm Me
        rgsPosicionarAjuda Me, fintForAtu, fbooForLog

        Set fclsContro = New clssContro
        Set fclsAceFun = New clssAceFun
        Set fclsAceFor = New clssAceFor
        Set fclsAgeCad = New clssAgeCad

        rlsCarregarAgencias

        rlsHabilitarBotoes

        fcmbF06Loc.Enabled = IIf(formMDIAce.menuConUsu.Enabled, True, False)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        gintTempoP = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
        Set fclsContro = Nothing
        Set fclsAceFun = Nothing
        Set fclsAceFor = Nothing
        Set fclsAgeCad = Nothing
        formMDIAce.fsbrModAce.Panels(4).Picture = LoadPicture()
        formMDIAce.ftbrModAce.Buttons("Usuari").Value = tbrUnpressed
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
Private Sub rlsCarregarAgencias()
        Dim lRStAgeCad As Recordset

        Set lRStAgeCad = fclsAgeCad.ConsultarTodas

            fcboCodAge.Clear
        Do _
            While (Not lRStAgeCad.EOF)
            fcboCodAge.AddItem lRStAgeCad!Codigo & " - " & lRStAgeCad!NomAge
            lRStAgeCad.MoveNext
        Loop
        lRStAgeCad.Close
End Sub
Private Sub rlsCarregarFuncoesDeAgencia()
        Dim lstrFuncao(1 To 3) As String

            lstrFuncao(1) = "Contador"
            lstrFuncao(2) = "Gerente"
            lstrFuncao(3) = "Operador"

            fcboFuncao.Clear

        For fbytIndice = 1 To 3
            fcboFuncao.AddItem lstrFuncao(fbytIndice)
        Next
End Sub
Private Sub rlsCarregarFuncoesDeDirecao()
        Dim lstrFuncao(1 To 5) As String

            lstrFuncao(1) = "Analista"
            lstrFuncao(2) = "Auditor"
            lstrFuncao(3) = "Contador"
            lstrFuncao(4) = "Gestor"
            lstrFuncao(5) = "Operador"

            fcboFuncao.Clear

        For fbytIndice = 1 To 5
            fcboFuncao.AddItem lstrFuncao(fbytIndice)
        Next
End Sub
Private Sub rlsConsultar()
        Dim lRStUsuari As Recordset

        Set lRStUsuari = gclsUsuari.Consultar(fintNumero)

        If (lRStUsuari.EOF) Then
            rlsLimparCampos
        Else
            fchkTemLog = IIf( _
                         lRStUsuari!TemLog, 1, 0)
            fcboStatus. _
             ListIndex = IIf( _
                         lRStUsuari!Status, 1, 0)
            ftxtNomUsu = lRStUsuari!NomUsu
            fintCodAge = lRStUsuari!CodAge

            rgsPesquisarComboIni fcboCodAge, fintCodAge

            fmskDatVal = Format(lRStUsuari!DatVal, "dd/mm/yyyy")

        If (fintCodAge = 0) Then
            rlsCarregarFuncoesDeDirecao
        Else
            rlsCarregarFuncoesDeAgencia
        End If

            rgsPesquisarComboAll fcboFuncao, lRStUsuari!Funcao

            ftxtE_Mail = lRStUsuari!E_Mail

            fcmbF02Inc.Enabled = False
            fcmbF03Alt.Enabled = fbooF03Alt
            fcmbF05Exc.Enabled = IIf(gclsAceMod.UsuarioAcessa(fintNumero) Or fclsAceFun.UsuarioAcessa(fintNumero) Or _
                                     fclsAceFor.UsuarioAcessa(fintNumero) Or gclsAceBot.UsuarioAcessa(fintNumero), False, fbooF05Exc)
        End If
        lRStUsuari.Close
End Sub
Private Sub rlsConsultarControle()
        Dim lRStContro As Recordset

        Set lRStContro = fclsContro.Consultar

            fintNumero = lRStContro!UltUsu + 1

        lRStContro.Close
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
        fchkTemLog = 0
        fcboStatus = ""
        ftxtNomUsu = ""
        fcboCodAge = ""
        fmskDatVal = ""
        fcboFuncao = ""
        ftxtE_Mail = ""
        fcmbF02Inc.Enabled = fbooF02Inc
        fcmbF03Alt.Enabled = False
        fcmbF05Exc.Enabled = False
End Sub
