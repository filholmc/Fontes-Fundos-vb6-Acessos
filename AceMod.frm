VERSION 5.00
Begin VB.Form formAceMod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acesso aos Módulos"
   ClientHeight    =   2430
   ClientLeft      =   4710
   ClientTop       =   1755
   ClientWidth     =   4770
   Icon            =   "AceMod.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "AceMod.frx":030A
   ScaleHeight     =   2430
   ScaleWidth      =   4770
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   0
      TabIndex        =   12
      Top             =   1860
      Width           =   4755
   End
   Begin VB.CommandButton fcmbEscape 
      Caption         =   "Fechar (Esc)"
      Height          =   315
      Left            =   3540
      TabIndex        =   5
      Top             =   2010
      Width           =   1120
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   510
      TabIndex        =   6
      Top             =   450
      Width           =   4245
   End
   Begin VB.ComboBox fcboModulo 
      Height          =   315
      ItemData        =   "AceMod.frx":0FD4
      Left            =   570
      List            =   "AceMod.frx":0FD6
      TabIndex        =   1
      Top             =   1440
      Width           =   3675
   End
   Begin VB.ComboBox fcboUsuari 
      Height          =   315
      ItemData        =   "AceMod.frx":0FD8
      Left            =   570
      List            =   "AceMod.frx":0FDA
      TabIndex        =   0
      Top             =   810
      Width           =   3675
   End
   Begin VB.CommandButton fcmbF09LCA 
      Caption         =   "F9"
      Height          =   255
      Left            =   3330
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF11Tab 
      Caption         =   "F11"
      Height          =   255
      Left            =   3780
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF12Hom 
      Caption         =   "F12"
      Height          =   255
      Left            =   4230
      TabIndex        =   9
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
      Caption         =   "Módulo"
      Height          =   195
      Left            =   570
      TabIndex        =   11
      Top             =   1230
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuário"
      Height          =   195
      Left            =   570
      TabIndex        =   10
      Top             =   600
      Width           =   540
   End
End
Attribute VB_Name = "formAceMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private fbooF02Inc As Boolean

Private fbooF05Exc As Boolean

Private fbooForLog As Boolean

Private fbytNumMod As Byte

Private fintForAtu As Integer

Private fintNumUsu As Integer
Private Sub fcboModulo_Change()
        If (fcboModulo = "") Then rlsDesabilitarBotoes
End Sub
Private Sub fcboModulo_Click()
        If (fcboModulo.ListIndex = -1) Then Exit Sub

        fbytNumMod = CByte(Mid(fcboModulo, 1, InStr(fcboModulo, " ") - 1))

        If (fcboUsuari.ListIndex = -1) Then Exit Sub

        rlsConsultar
End Sub
Private Sub fcboUsuari_Change()
        If (fcboUsuari = "") Then rlsDesabilitarBotoes
End Sub
Private Sub fcboUsuari_Click()
        If (fcboUsuari.ListIndex = -1) Then Exit Sub

        fintNumUsu = CInt(Mid(fcboUsuari, 1, InStr(fcboUsuari, " ") - 1))

        If (fcboModulo.ListIndex = -1) Then Exit Sub

        rlsConsultar
End Sub
Private Sub fcmbEscape_Click()
        Unload Me
End Sub
Private Sub fcmbF02Inc_Click()
            gclsAceMod.Incluir fintNumUsu, fbytNumMod
            gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                        "Incluiu", 0, Format(fintNumUsu, "0") & "; " & _
                                                                      Format(fbytNumMod, "0"), " "
        If (gDBCFundos.Errors.Count > 0) Then
            rgsTratarErro Err, gDBCFundos.Errors, Me
        Else
            rgfMsgBox "Acesso Incluído", MsgInf
            fcmbF02Inc.Enabled = False
            fcmbF05Exc.Enabled = fbooF05Exc
            fcboModulo.SetFocus
        End If
End Sub
Private Sub fcmbF05Exc_Click()
        If (rgfMsgBox("Confirma Exclusão?", MsgNao) = vbYes) Then
            gclsAceMod.Excluir fintNumUsu, fbytNumMod
            gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                        "Excluiu", 0, Format(fintNumUsu, "0") & "; " & _
                                                                      Format(fbytNumMod, "0"), " "
        If (gDBCFundos.Errors.Count > 0) Then
            rgsTratarErro Err, gDBCFundos.Errors, Me
        Else
            rgfMsgBox "Acesso Excluído", MsgInf
            fcmbF02Inc.Enabled = fbooF02Inc
            fcmbF05Exc.Enabled = False
        End If
        End If
        fcboModulo.SetFocus
End Sub
Private Sub fcmbF06Loc_Click()
        Select Case gbytConMod
               Case 1
                    formConAMo.SetFocus
               Case 2
                    formConMoA.SetFocus
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
        If (Not gbooConMod) Then Exit Sub

        gbooConMod = False
        rgsPesquisarComboIni fcboUsuari, gintNumUsu
        rgsPesquisarComboIni fcboModulo, gbytNumMod
        rlsConsultar
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        gintTempoP = 0
        rgsTratarFuncoes KeyCode, Me
End Sub
Private Sub Form_Load()
        formMDIAce.ftbrModAce.Buttons("AceMod").Value = tbrPressed

        rgsCentralizarForm Me
        rgsPosicionarAjuda Me, fintForAtu, fbooForLog

        rgsCarregarUsuarios fcboUsuari, False

        rgsCarregarModulos fcboModulo

        rlsHabilitarBotoes

                gbytConMod = IIf((gbytConMod < 2 And _
                                                 Not formMDIAce.menuConAMo.Enabled) And _
                                                     formMDIAce.menuConMoA.Enabled, 2, 1)
        fcmbF06Loc.Enabled = IIf((gbytConMod < 2 And formMDIAce.menuConAMo.Enabled) Or _
                                 (gbytConMod = 2 And formMDIAce.menuConMoA.Enabled), True, False)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        gintTempoP = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
        formMDIAce.fsbrModAce.Panels(4).Picture = LoadPicture()
        formMDIAce.ftbrModAce.Buttons("AceMod").Value = tbrUnpressed
End Sub
Private Sub rlsConsultar()
        Dim lRStAceMod As Recordset

        Set lRStAceMod = gclsAceMod.Consultar(fintNumUsu, fbytNumMod)

        If (fbytNumMod > 0) Then
        If (lRStAceMod.EOF) Then
            fcmbF02Inc.Enabled = fbooF02Inc
            fcmbF05Exc.Enabled = False
        Else
            fcmbF02Inc.Enabled = False
            fcmbF05Exc.Enabled = fbooF05Exc
        End If
        End If
        lRStAceMod.Close
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
