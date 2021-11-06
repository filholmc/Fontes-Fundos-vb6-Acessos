VERSION 5.00
Begin VB.Form formModulo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulos"
   ClientHeight    =   2670
   ClientLeft      =   4200
   ClientTop       =   1755
   ClientWidth     =   5910
   Icon            =   "Modulo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "Modulo.frx":030A
   ScaleHeight     =   2670
   ScaleWidth      =   5910
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   0
      TabIndex        =   13
      Top             =   2100
      Width           =   5895
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   510
      TabIndex        =   12
      Top             =   450
      Width           =   5385
   End
   Begin VB.TextBox ftxtDescri 
      Height          =   315
      Left            =   2194
      LinkTimeout     =   0
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1455
      Width           =   2235
   End
   Begin VB.TextBox ftxtNumero 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2194
      MaxLength       =   2
      MultiLine       =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Na Inclusão, o Número é gerado automaticamente"
      Top             =   735
      Width           =   315
   End
   Begin VB.CommandButton fcmbF09LCA 
      Caption         =   "F9"
      Height          =   255
      Left            =   4470
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF11Tab 
      Caption         =   "F11"
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF12Hom 
      Caption         =   "F12"
      Height          =   255
      Left            =   5370
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF06Loc 
      Caption         =   "Localizar (F6)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3540
      TabIndex        =   5
      Top             =   2250
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF02Inc 
      Caption         =   "Incluir (F2)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2250
      Width           =   1120
   End
   Begin VB.CommandButton fcmbEscape 
      Caption         =   "Fechar (Esc)"
      Height          =   315
      Left            =   4680
      TabIndex        =   6
      Top             =   2250
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF03Alt 
      Caption         =   "Alterar (F3)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1260
      TabIndex        =   3
      Top             =   2250
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF05Exc 
      Caption         =   "Excluir (F5)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Top             =   2250
      Width           =   1120
   End
   Begin VB.Label flblDescri 
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
      Height          =   195
      Index           =   1
      Left            =   1421
      TabIndex        =   11
      Top             =   1485
      Width           =   720
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
      Index           =   0
      Left            =   1421
      LinkTimeout     =   0
      TabIndex        =   10
      Top             =   765
      Width           =   660
   End
End
Attribute VB_Name = "formModulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private fbooF02Inc As Boolean

Private fbooF03Alt As Boolean

Private fbooF05Exc As Boolean

Private fbooForLog As Boolean

Private fbytNumero As Byte

Private fclsContro As clssContro

Private fintForAtu As Integer
Private Function rlfValidacaoDeCamposOk() As Boolean
        rlfValidacaoDeCamposOk = False

        If (gclsModulo.DescricaoCadastrada(ftxtDescri)) Then
            rgfMsgBox "Descrição já utilizada", MsgErr
            ftxtDescri.SetFocus
            Exit Function
        End If

        If (Trim(ftxtDescri) = "") Then
            rgfMsgBox "Preencha o campo 'Descrição'", MsgErr, Me.HelpContextID
            ftxtDescri.SetFocus
            Exit Function
        End If

        rlfValidacaoDeCamposOk = True
End Function
Private Sub fcmbEscape_Click()
        Unload Me
End Sub
Private Sub fcmbF02Inc_Click()
        If (Not rlfValidacaoDeCamposOk) Then Exit Sub

        rlsConsultarControle

        gDBCFundos.BeginTrans
                   fclsContro.Alterar "UltMod", fbytNumero

               If (gDBCFundos.Errors.Count > 0) Then GoTo Erro_DB

                   gclsModulo.Incluir fbytNumero, ftxtDescri

               If (gDBCFundos.Errors.Count > 0) Then GoTo Erro_DB

                   gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                               "Incluiu", 0, Format(fbytNumero, "0"), ftxtDescri
               If (gDBCFundos.Errors.Count > 0) Then GoTo Erro_DB
        gDBCFundos.CommitTrans

        rgfMsgBox "Módulo Incluído", MsgInf

        fcmbF02Inc.Enabled = False
        fcmbF03Alt.Enabled = fbooF03Alt
        fcmbF05Exc.Enabled = fbooF05Exc
        ftxtNumero = fbytNumero
        ftxtNumero.SetFocus
        Exit Sub

Erro_DB:
        gDBCFundos.RollbackTrans

        rgsTratarErro Err, gDBCFundos.Errors, Me
End Sub
Private Sub fcmbF03Alt_Click()
        If (Not rlfValidacaoDeCamposOk) Then Exit Sub

            gclsModulo.Alterar fbytNumero, ftxtDescri
            gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                        "Alterou", 0, Format(fbytNumero, "0"), ftxtDescri
        If (gDBCFundos.Errors.Count > 0) Then
            rgsTratarErro Err, gDBCFundos.Errors, Me
        Else
            rgfMsgBox "Descrição do Módulo alterada", MsgInf
            ftxtNumero.SetFocus
        End If
End Sub
Private Sub fcmbF05Exc_Click()
        If (rgfMsgBox("Confirma Exclusão?", MsgNao) = vbYes) Then
            gclsModulo.Excluir (fbytNumero)
            gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                        "Excluiu", 0, Format(fbytNumero, "0"), " "
        If (gDBCFundos.Errors.Count > 0) Then
            rgsTratarErro Err, gDBCFundos.Errors, Me
        Else
            rgfMsgBox "Módulo Excluído", MsgInf
            fcmbF03Alt.Enabled = False
            fcmbF05Exc.Enabled = False
        End If
        End If
        ftxtNumero.SetFocus
End Sub
Private Sub fcmbF06Loc_Click()
        formConMod.SetFocus
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
        If (Not gbooConMod) Then Exit Sub

        gbooConMod = False
        ftxtNumero = gbytNumMod
        ftxtNumero_LostFocus
        ftxtNumero.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        gintTempoP = 0
        rgsTratarFuncoes KeyCode, Me
End Sub
Private Sub Form_Load()
        formMDIAce.ftbrModAce.Buttons("Modulo").Value = tbrPressed

        rgsCentralizarForm Me
        rgsPosicionarAjuda Me, fintForAtu, fbooForLog

        Set fclsContro = New clssContro

        rlsHabilitarBotoes

        fcmbF06Loc.Enabled = IIf(formMDIAce.menuConMod.Enabled, True, False)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        gintTempoP = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
        Set fclsContro = Nothing
        formMDIAce.fsbrModAce.Panels(4).Picture = LoadPicture()
        formMDIAce.ftbrModAce.Buttons("Modulo").Value = tbrUnpressed
End Sub
Private Sub ftxtNumero_GotFocus()
        rlsDesabilitarBotoes
        ftxtNumero.SelStart = Len(ftxtNumero)
End Sub
Private Sub ftxtNumero_KeyPress(KeyAscii As Integer)
        If (Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8) Then KeyAscii = 0
End Sub
Private Sub ftxtNumero_LostFocus()
        rlsFormatarChaves
        rlsConsultar
End Sub
Private Sub rlsConsultar()
        Dim lRStModulo As Recordset

        Set lRStModulo = gclsModulo.Consultar(fbytNumero)

        If (lRStModulo.EOF) Then
            rlsLimparCampos
        Else
            ftxtDescri = lRStModulo!Descri

            fcmbF02Inc.Enabled = False
            fcmbF03Alt.Enabled = fbooF03Alt
            fcmbF05Exc.Enabled = IIf(gclsFormes.PertenceModulo(fbytNumero) Or gclsAceMod.ModuloAcessado(fbytNumero), False, fbooF05Exc)
        End If
        lRStModulo.Close
End Sub
Private Sub rlsConsultarControle()
        Dim lRStContro As Recordset

        Set lRStContro = fclsContro.Consultar

            fbytNumero = lRStContro!UltMod + 1

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

        fbytNumero = CByte(ftxtNumero)
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
        ftxtDescri = ""
        fcmbF02Inc.Enabled = fbooF02Inc
        fcmbF03Alt.Enabled = False
        fcmbF05Exc.Enabled = False
End Sub
