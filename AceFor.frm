VERSION 5.00
Begin VB.Form formAceFor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acesso aos Forms"
   ClientHeight    =   2730
   ClientLeft      =   4770
   ClientTop       =   1755
   ClientWidth     =   4770
   Icon            =   "AceFor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "AceFor.frx":030A
   ScaleHeight     =   2730
   ScaleWidth      =   4770
   Begin VB.CheckBox fchkAceBot 
      Caption         =   "Dar Acesso a Todos os Botões do Módulo"
      Enabled         =   0   'False
      Height          =   195
      Left            =   570
      TabIndex        =   3
      Top             =   1860
      Width           =   3315
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   0
      TabIndex        =   15
      Top             =   2160
      Width           =   4755
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   510
      TabIndex        =   14
      Top             =   450
      Width           =   4245
   End
   Begin VB.ComboBox fcboModulo 
      Height          =   315
      ItemData        =   "AceFor.frx":0BD4
      Left            =   570
      List            =   "AceFor.frx":0BD6
      TabIndex        =   1
      Top             =   1440
      Width           =   1800
   End
   Begin VB.ComboBox fcboUsuari 
      Height          =   315
      ItemData        =   "AceFor.frx":0BD8
      Left            =   570
      List            =   "AceFor.frx":0BDA
      TabIndex        =   0
      Top             =   810
      Width           =   3705
   End
   Begin VB.ComboBox fcboFormes 
      Height          =   315
      ItemData        =   "AceFor.frx":0BDC
      Left            =   2490
      List            =   "AceFor.frx":0BDE
      TabIndex        =   2
      Top             =   1440
      Width           =   1800
   End
   Begin VB.CommandButton fcmbF09LCA 
      Caption         =   "F9"
      Height          =   255
      Left            =   3330
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF11Tab 
      Caption         =   "F11"
      Height          =   255
      Left            =   3780
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF12Hom 
      Caption         =   "F12"
      Height          =   255
      Left            =   4230
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF06Loc 
      Caption         =   "Localizar (F6)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2400
      TabIndex        =   6
      Top             =   2310
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF05Exc 
      Caption         =   "Excluir (F5)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1260
      TabIndex        =   5
      Top             =   2310
      Width           =   1120
   End
   Begin VB.CommandButton fcmbEscape 
      Caption         =   "Fechar (Esc)"
      Height          =   315
      Left            =   3540
      TabIndex        =   7
      Top             =   2310
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF02Inc 
      Caption         =   "Incluir (F2)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   2310
      Width           =   1120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Módulo"
      Height          =   195
      Left            =   570
      TabIndex        =   13
      Top             =   1230
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuário"
      Height          =   195
      Left            =   570
      TabIndex        =   12
      Top             =   600
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Form"
      Height          =   195
      Left            =   2490
      TabIndex        =   11
      Top             =   1230
      Width           =   345
   End
End
Attribute VB_Name = "formAceFor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private fbooF02Inc As Boolean

Private fbooF05Exc As Boolean

Private fbooForLog As Boolean

Private fbytNumMod As Byte

Private fbytQtdAcB As Byte

Private fclsBotoes As clssBotoes

Private fclsAceFor As clssAceFor

Private fintForAtu As Integer

Private fintNumFor As Integer, fintNumUsu As Integer
Private Sub fcboFormes_Change()
        If (fcboFormes = "") Then rlsDesabilitarBotoes
End Sub
Private Sub fcboFormes_Click()
        If (fcboFormes.ListIndex = -1) Then Exit Sub

        fintNumFor = CInt(Mid(fcboFormes, InStr(fcboFormes, "-") + 2, Len(fcboFormes) - InStr(fcboFormes, "-") + 1))
        rlsConsultar
        fchkAceBot = 0
        fchkAceBot.Enabled = IIf(fintNumFor = 0, True, False)
End Sub
Private Sub fcboModulo_Change()
        If (fcboModulo = "") Then rlsDesabilitarBotoes
End Sub
Private Sub fcboModulo_Click()
        If (fcboModulo.ListIndex = -1) Then Exit Sub

        fbytNumMod = CByte(Mid(fcboModulo, 1, InStr(fcboModulo, " ") - 1))

        rlsDesabilitarBotoes

        rgsCarregarFormesDeUmModuloAcessaveis fbytNumMod, fcboFormes, True
End Sub
Private Sub fcboUsuari_Change()
        If (fcboUsuari = "") Then rlsDesabilitarBotoes
End Sub
Private Sub fcboUsuari_Click()
        If (fcboUsuari.ListIndex = -1) Then Exit Sub

        fintNumUsu = CInt(Mid(fcboUsuari, 1, InStr(fcboUsuari, " ") - 1))

        rlsDesabilitarBotoes

        rgsCarregarModulosDeUmUsuario fintNumUsu, fcboModulo
End Sub
Private Sub fcmbEscape_Click()
        Unload Me
End Sub
Private Sub fcmbF02Inc_Click()
            rlsIncluir

        If (gDBCFundos.Errors.Count > 0) Then
            rgsTratarErro Err, gDBCFundos.Errors, Me
        Else
            rgfMsgBox IIf(fintNumFor > 0, "Acesso Incluído", gbytQtdAce & " Acessos a Forms" & _
                      IIf(fchkAceBot = 1, " e " & fbytQtdAcB & " Acessos a Botões", "") & " Incluídos"), MsgInf
            fcmbF02Inc.Enabled = False
            fcmbF05Exc.Enabled = fbooF05Exc
            fcboFormes.SetFocus
        End If
End Sub
Private Sub fcmbF05Exc_Click()
        If (rgfMsgBox("Confirma Exclusão?", MsgNao) = vbYes) Then
            fclsAceFor.Excluir fintNumUsu, fintNumFor
            gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                        "Excluiu", 0, Format(fintNumUsu, "0") & "; " & _
                                                                      Format(fintNumFor, "0"), " "
        If (gDBCFundos.Errors.Count > 0) Then
            rgsTratarErro Err, gDBCFundos.Errors, Me
        Else
            rgfMsgBox "Acesso Excluído", MsgInf
            fcmbF02Inc.Enabled = fbooF02Inc
            fcmbF05Exc.Enabled = False
        End If
        End If
        fcboFormes.SetFocus
End Sub
Private Sub fcmbF06Loc_Click()
        Select Case gbytConFor
               Case 1
                    formConAFo.SetFocus
               Case 2
                    formConFoA.SetFocus
        End Select
End Sub
Private Sub fcmbF09LCA_Click()
        If (Not TypeOf ActiveControl Is CommandButton) Then ActiveControl.Text = ""
End Sub
Private Sub fcmbF11Tab_Click()
        SendKeys "+{TAB}"
End Sub
Private Sub fcmbF12Hom_Click()
        fcboUsuari.SetFocus
End Sub
Private Sub Form_Activate()
        gintTempoP = 0
        formMDIAce.fsbrModAce.Panels(4).Picture = IIf(gbooUsuLog Or fbooForLog, _
                                                      formMDIAce.fimlStaBar.ListImages(1).Picture, LoadPicture())
        If (Not gbooConFor) Then Exit Sub

        gbooConFor = False
        rgsPesquisarComboIni fcboUsuari, gintNumUsu
        rgsPesquisarComboIni fcboModulo, gbytNumMod
        rgsPesquisarComboFim fcboFormes, gintNumFor
        rlsConsultar
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        gintTempoP = 0
        rgsTratarFuncoes KeyCode, Me
End Sub
Private Sub Form_Load()
        formMDIAce.ftbrModAce.Buttons("AceFor").Value = tbrPressed

        rgsCentralizarForm Me
        rgsPosicionarAjuda Me, fintForAtu, fbooForLog

        Set fclsBotoes = New clssBotoes
        Set fclsAceFor = New clssAceFor

        rgsCarregarUsuarios fcboUsuari, False

        rlsHabilitarBotoes

                gbytConFor = IIf((gbytConFor < 2 And _
                                                 Not formMDIAce.menuConAFo.Enabled) And _
                                                     formMDIAce.menuConFoA.Enabled, 2, 1)
        fcmbF06Loc.Enabled = IIf((gbytConFor < 2 And formMDIAce.menuConAFo.Enabled) Or _
                                 (gbytConFor = 2 And formMDIAce.menuConFoA.Enabled), True, False)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        gintTempoP = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
        Set fclsBotoes = Nothing
        Set fclsAceFor = Nothing
        formMDIAce.fsbrModAce.Panels(4).Picture = LoadPicture()
        formMDIAce.ftbrModAce.Buttons("AceFor").Value = tbrUnpressed
End Sub
Private Sub rlsConsultar()
        Dim lRStAceFor As Recordset

        Set lRStAceFor = fclsAceFor.Consultar(fintNumUsu, fintNumFor)

        If (lRStAceFor.EOF) Then
            fcmbF02Inc.Enabled = fbooF02Inc
            fcmbF05Exc.Enabled = False
        Else
            fcmbF02Inc.Enabled = False
            fcmbF05Exc.Enabled = fbooF05Exc
        End If
        lRStAceFor.Close
End Sub
Private Sub rlsDesabilitarBotoes()
        fcmbF02Inc.Enabled = False
        fcmbF05Exc.Enabled = False
End Sub
Private Sub rlsHabilitarBotoes()
        Dim lRStAceBot As Recordset

        Set lRStAceBot = gclsAceBot.ConsultarBotoesDeUmUsuarioPorModuloAndForm(gintUsuLog, 1, fintForAtu)

        Do _
            While (Not ((lRStAceBot.EOF)))
            Select Case (lRStAceBot!NomBot)
                   Case "fcmbF02Inc"
                         fbooF02Inc = True
                   Case "fcmbF05Exc"
                         fbooF05Exc = True
            End Select
            lRStAceBot.MoveNext
        Loop
        lRStAceBot.Close
End Sub
Private Sub rlsIncluir()
        Dim lRStFormes As Recordset

        Set lRStFormes = gclsFormes.ConsultarFormsDeUmModuloAcessaveis(fbytNumMod)

            gbytQtdAce = 0

        If (fintNumFor > 0) Then
            fclsAceFor.Incluir fintNumUsu, fintNumFor
            gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                        "Incluiu", 0, Format(fintNumUsu, "0") & "; " & _
                                                                      Format(fintNumFor, "0"), " "
        Else
            gstrCteudo = "Todos os Forms"

        If (fchkAceBot = 1) Then rlsIncluirBotoes

        If (gDBCFundos.Errors.Count > 0) Then Exit Sub
        Do _
            While (Not lRStFormes.EOF)
        If (fclsAceFor.Ausente(fintNumUsu, lRStFormes!Numero)) Then
            fclsAceFor.Incluir fintNumUsu, lRStFormes!Numero

            gbytQtdAce = gbytQtdAce + 1
        If (gDBCFundos.Errors.Count > 0) Then Exit Sub
        End If
            lRStFormes.MoveNext
        Loop
            gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                        "Incluiu", 0, Format(fintNumUsu, "0") & "; " & _
                                                                      Format(fbytNumMod, "0"), gstrCteudo
        End If
        lRStFormes.Close
End Sub
Private Sub rlsIncluirBotoes()
        Dim lRStBotoes As Recordset

        Set lRStBotoes = fclsBotoes.ConsultarBotoesDeUmModulo(fbytNumMod)

            fbytQtdAcB = 0
            gstrCteudo = gstrCteudo & " e Botões"
        Do _
            While (Not lRStBotoes.EOF)
        If (gclsAceBot.Ausente(fintNumUsu, lRStBotoes!Numero)) Then
            gclsAceBot.Incluir fintNumUsu, lRStBotoes!Numero

            fbytQtdAcB = fbytQtdAcB + 1
        If (gDBCFundos.Errors.Count > 0) Then Exit Sub
        End If
            lRStBotoes.MoveNext
        Loop
        lRStBotoes.Close
End Sub
