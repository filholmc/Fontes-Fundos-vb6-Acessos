VERSION 5.00
Begin VB.Form formAceFun 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acesso aos Fundos"
   ClientHeight    =   2430
   ClientLeft      =   4680
   ClientTop       =   1755
   ClientWidth     =   4770
   Icon            =   "AceFun.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "AceFun.frx":030A
   ScaleHeight     =   2430
   ScaleWidth      =   4770
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   0
      TabIndex        =   12
      Top             =   1860
      Width           =   4755
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   510
      TabIndex        =   11
      Top             =   450
      Width           =   4245
   End
   Begin VB.ComboBox fcboFundos 
      Height          =   315
      ItemData        =   "AceFun.frx":0BD4
      Left            =   570
      List            =   "AceFun.frx":0BD6
      TabIndex        =   1
      Top             =   1440
      Width           =   3645
   End
   Begin VB.ComboBox fcboUsuari 
      Height          =   315
      ItemData        =   "AceFun.frx":0BD8
      Left            =   570
      List            =   "AceFun.frx":0BDA
      TabIndex        =   0
      Top             =   810
      Width           =   3645
   End
   Begin VB.CommandButton fcmbF09LCA 
      Caption         =   "F9"
      Height          =   255
      Left            =   3330
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF11Tab 
      Caption         =   "F11"
      Height          =   255
      Left            =   3780
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF12Hom 
      Caption         =   "F12"
      Height          =   255
      Left            =   4230
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF06Loc 
      Caption         =   "Localizar (F6)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Top             =   2010
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF05Exc 
      Caption         =   "Excluir (F5)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1260
      TabIndex        =   3
      Top             =   2010
      Width           =   1120
   End
   Begin VB.CommandButton fcmbEscape 
      Caption         =   "Fechar (Esc)"
      Height          =   315
      Left            =   3540
      TabIndex        =   5
      Top             =   2010
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF02Inc 
      Caption         =   "Incluir (F2)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2010
      Width           =   1120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fundo"
      Height          =   195
      Left            =   570
      TabIndex        =   10
      Top             =   1230
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuário"
      Height          =   195
      Left            =   570
      TabIndex        =   9
      Top             =   600
      Width           =   540
   End
End
Attribute VB_Name = "formAceFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private fbooF02Inc As Boolean

Private fbooF05Exc As Boolean

Private fbooForLog As Boolean

Private fbytNumFun As Byte

Private fclsAceFun As clssAceFun

Private fintForAtu As Integer

Private fintNumUsu As Integer
Private Sub fcboFundos_Change()
        If (fcboFundos = "") Then rlsDesabilitarBotoes
End Sub
Private Sub fcboFundos_Click()
        If (fcboFundos.ListIndex = -1) Then Exit Sub

        fbytNumFun = CByte(Mid(fcboFundos, InStrRev(fcboFundos, "-") + 2, Len(fcboFundos) - InStrRev(fcboFundos, "-") + 1))

        If (fcboUsuari.ListIndex = -1) Then Exit Sub

        rlsConsultar
End Sub
Private Sub fcboUsuari_Change()
        If (fcboUsuari = "") Then rlsDesabilitarBotoes
End Sub
Private Sub fcboUsuari_Click()
        If (fcboUsuari.ListIndex = -1) Then Exit Sub

        fintNumUsu = CInt(Mid(fcboUsuari, 1, InStr(fcboUsuari, " ") - 1))

        If (fcboFundos.ListIndex = -1) Then Exit Sub

        rlsConsultar
End Sub
Private Sub fcmbEscape_Click()
        Unload Me
End Sub
Private Sub fcmbF02Inc_Click()
            rlsIncluir

        If (gDBCFundos.Errors.Count > 0) Then
            rgsTratarErro Err, gDBCFundos.Errors, Me
        Else
            rgfMsgBox IIf(fintNumUsu > 0, "Acesso Incluído", gbytQtdAce & " Acesso(s) Incluído(s)"), MsgInf
            fcmbF02Inc.Enabled = False
            fcmbF05Exc.Enabled = True
            fcboFundos.SetFocus
        End If
End Sub
Private Sub fcmbF05Exc_Click()
        If (rgfMsgBox("Confirma Exclusão?", MsgNao) = vbYes) Then
            fclsAceFun.Excluir fintNumUsu, fbytNumFun
            gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                        "Excluiu", 0, Format(fintNumUsu, "0") & "; " & _
                                                                      Format(fbytNumFun, "0"), " "
        If (gDBCFundos.Errors.Count > 0) Then
            rgsTratarErro Err, gDBCFundos.Errors, Me
        Else
            rgfMsgBox "Acesso Excluído", MsgInf
            fcmbF02Inc.Enabled = fbooF02Inc
            fcmbF05Exc.Enabled = False
        End If
        End If
        fcboFundos.SetFocus
End Sub
Private Sub fcmbF06Loc_Click()
        Select Case gbytConFun
               Case 1
                    formConAFu.SetFocus
               Case 2
                    formConFuA.SetFocus
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
        If (Not gbooConFun) Then Exit Sub

        gbooConFun = False
        rgsPesquisarComboIni fcboUsuari, gintNumUsu
        rgsPesquisarComboFim fcboFundos, gbytNumFun
        rlsConsultar
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        gintTempoP = 0
        rgsTratarFuncoes KeyCode, Me
End Sub
Private Sub Form_Load()
        formMDIAce.ftbrModAce.Buttons("AceFun").Value = tbrPressed

        rgsCentralizarForm Me
        rgsPosicionarAjuda Me, fintForAtu, fbooForLog

        Set fclsAceFun = New clssAceFun

        rgsCarregarUsuarios fcboUsuari, True

        rgsCarregarFundos fcboFundos

        rlsHabilitarBotoes

                gbytConFun = IIf((gbytConFun < 2 And _
                                                 Not formMDIAce.menuConAFu.Enabled) And _
                                                     formMDIAce.menuConFuA.Enabled, 2, 1)
        fcmbF06Loc.Enabled = IIf((gbytConFun < 2 And formMDIAce.menuConAFu.Enabled) Or _
                                 (gbytConFun = 2 And formMDIAce.menuConFuA.Enabled), True, False)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        gintTempoP = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
        Set fclsAceFun = Nothing
        formMDIAce.fsbrModAce.Panels(4).Picture = LoadPicture()
        formMDIAce.ftbrModAce.Buttons("AceFun").Value = tbrUnpressed
End Sub
Private Sub rlsConsultar()
        Dim lRStAceFun As Recordset

        Set lRStAceFun = fclsAceFun.Consultar(fintNumUsu, fbytNumFun)

        If (fbytNumFun > 0) Then
        If (lRStAceFun.EOF) Then
            fcmbF02Inc.Enabled = fbooF02Inc
            fcmbF05Exc.Enabled = False
        Else
            fcmbF02Inc.Enabled = False
            fcmbF05Exc.Enabled = fbooF05Exc
        End If
        End If
        lRStAceFun.Close
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
        Dim lRStUsuari As Recordset

        Set lRStUsuari = gclsUsuari.ConsultarTodosPorNumero

            gbytQtdAce = 0

        If (fintNumUsu > 0) Then
            fclsAceFun.Incluir fintNumUsu, fbytNumFun
            gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                        "Incluiu", 0, Format(fintNumUsu, "0") & "; " & _
                                                                      Format(fbytNumFun, "0"), " "
        Else
        Do _
            While (Not lRStUsuari.EOF)
        If (fclsAceFun.Ausente(lRStUsuari!Numero, fbytNumFun)) Then
            fclsAceFun.Incluir lRStUsuari!Numero, fbytNumFun

            gbytQtdAce = gbytQtdAce + 1
        If (gDBCFundos.Errors.Count > 0) Then Exit Sub
        End If
            lRStUsuari.MoveNext
        Loop
            gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                        "Incluiu", 0, Format(fbytNumFun, "0"), "Todos os Usuários"
        End If
        lRStUsuari.Close
End Sub
