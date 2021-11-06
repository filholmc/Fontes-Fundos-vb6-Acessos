VERSION 5.00
Begin VB.Form formTroUsu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Troca de Usuário"
   ClientHeight    =   3240
   ClientLeft      =   4020
   ClientTop       =   1755
   ClientWidth     =   6165
   Icon            =   "TroUsu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6165
   Begin VB.CommandButton fcmbF12Hom 
      Caption         =   "F12"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   390
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF11Tab 
      Caption         =   "F11"
      Height          =   255
      Left            =   2430
      TabIndex        =   11
      Top             =   390
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF09LCA 
      Caption         =   "F9"
      Height          =   255
      Left            =   1980
      TabIndex        =   10
      Top             =   390
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Frame frmeFrame2 
      Caption         =   "Escolha um outro Usuário para Operrar:"
      Height          =   915
      Left            =   1980
      TabIndex        =   0
      Top             =   1650
      Width           =   3945
      Begin VB.TextBox ftxtSenhas 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2370
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox ftxtNumUsu 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   810
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   555
      End
      Begin VB.Label flblSenhas 
         AutoSize        =   -1  'True
         Caption         =   "Senha"
         Height          =   195
         Left            =   1800
         LinkTimeout     =   0
         TabIndex        =   9
         Top             =   420
         Width           =   465
      End
      Begin VB.Label flblNumero 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   150
         LinkTimeout     =   0
         TabIndex        =   8
         Top             =   420
         Width           =   555
      End
   End
   Begin VB.CommandButton fcmbEscape 
      Caption         =   "Fechar (Esc)"
      Height          =   315
      Left            =   4800
      TabIndex        =   4
      Top             =   2820
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF03Tro 
      Caption         =   "Trocar (F3)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   315
      Left            =   3660
      TabIndex        =   3
      Top             =   2820
      Width           =   1120
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1980
      MultiLine       =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "TroUsu.frx":030A
      Top             =   660
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   240
      TabIndex        =   5
      Top             =   2670
      Width           =   5715
   End
   Begin VB.Label flblNomUsu 
      Height          =   195
      Left            =   1980
      TabIndex        =   7
      Top             =   900
      Width           =   3915
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2265
      Left            =   300
      Picture         =   "TroUsu.frx":032C
      Stretch         =   -1  'True
      Top             =   300
      Width           =   1590
   End
End
Attribute VB_Name = "formTroUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private fbooForLog As Boolean

Private fbooUsuLog As Boolean

Private fintForAtu As Integer

Private fintNumUsu As Integer

Private fstrNomUsu As String, fstrSenhas As String
Private Sub fcmbEscape_Click()
        Unload Me
End Sub
Private Sub fcmbF03Tro_Click()
        If (ftxtSenhas = "") Then
            rgfMsgBox "Preencha o campo 'Senha'", MsgErr, Me.HelpContextID
            ftxtSenhas.SetFocus
            Exit Sub
        End If

        If (fstrSenhas <> rgfSenhaCp(ftxtSenhas)) Then
            rgfMsgBox "Senha não confere", MsgErr
            ftxtSenhas.SetFocus
            Exit Sub
        End If

            gclsLogado.Alterar gstrNomCmp, fintNumUsu, gbytModAtv
            gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                         "Trocou", 0, Format(fintNumUsu, "0"), " "

        If (gDBCFundos.Errors.Count > 0) Then rgsTratarErro Err, gDBCFundos.Errors, Me

            gbooUsuLog = fbooUsuLog
            gintUsuLog = fintNumUsu
            gstrNomUsu = fstrNomUsu
            gstrSenhas = fstrSenhas
            formMDIAce.fsbrModAce.Panels(4) = gstrNomUsu
            Unload Me
End Sub
Private Sub fcmbF09LCA_Click()
        If (Not TypeOf ActiveControl Is CommandButton) Then ActiveControl.Text = ""
End Sub
Private Sub fcmbF11Tab_Click()
        SendKeys "+{TAB}"
End Sub
Private Sub fcmbF12Hom_Click()
        ftxtNumUsu.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        gintTempoP = 0
        rgsTratarFuncoes KeyCode, Me
End Sub
Private Sub Form_Load()
        gintTempoP = 0
        flblNomUsu = gintUsuLog & " - " & gstrNomUsu
        formMDIAce.fsbrModAce.Panels(4).Picture = IIf(gbooUsuLog Or fbooForLog, _
                                                      formMDIAce.fimlStaBar.ListImages(1).Picture, LoadPicture())
        rgsCentralizarFormIndependente Me
        rgsPosicionarAjuda Me, fintForAtu, fbooForLog
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        gintTempoP = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
        formMDIAce.fsbrModAce.Panels(4).Picture = LoadPicture()
End Sub
Private Sub ftxtNumUsu_GotFocus()
        rlsDesabilitarBotao
        ftxtNumUsu.SelStart = Len(ftxtNumUsu)
End Sub
Private Sub ftxtNumUsu_KeyPress(KeyAscii As Integer)
        If (Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8) Then KeyAscii = 0
End Sub
Private Sub ftxtNumUsu_KeyUp(KeyCode As Integer, Shift As Integer)
        If (KeyCode = 8) Or (KeyCode >= 48 And KeyCode <= 57) Or (KeyCode >= 96 And KeyCode <= 105) Then
            ftxtNumUsu = Format(ftxtNumUsu, "#,##0")
            ftxtNumUsu.SelStart = Len(ftxtNumUsu)
        End If
End Sub
Private Sub ftxtNumUsu_LostFocus()
        rlsFormatarChaves

        If (fintNumUsu = gintUsuLog) Then
            rgfMsgBox "Usuário já está realizando esta Sessão", MsgErr
            ftxtNumUsu.SetFocus
            Exit Sub
        End If
End Sub
Private Sub ftxtSenhas_GotFocus()
        rlsConsultar
End Sub
Private Sub rlsConsultar()
        Dim lRStUsuari As Recordset

        Set lRStUsuari = gclsUsuari.Consultar(fintNumUsu)

        If (lRStUsuari.EOF) Then
            rgfMsgBox "Usuário não cadastrado", MsgErr
            ftxtNumUsu.SetFocus
            rlsLimparCampos
            Exit Sub
        Else
            fstrSenhas = lRStUsuari!Senhas
            fbooUsuLog = IIf( _
                         lRStUsuari!TemLog, 1, 0)
            fstrNomUsu = lRStUsuari!NomUsu

        If (lRStUsuari!Status) Then
            rgfMsgBox "Usuário com Acesso Bloqueado", MsgErr
            ftxtNumUsu.SetFocus
            Exit Sub
        End If

        If (gdatServBD > lRStUsuari!DatVal) Then
            rgfMsgBox "Usuário com Acesso Expirado", MsgErr
            ftxtNumUsu.SetFocus
            Exit Sub
        End If

        If (gclsAceMod.Ausente(fintNumUsu, 1)) Then
            rgfMsgBox "Usuário não possui Acesso a este Módulo", MsgErr
            ftxtNumUsu.SetFocus
            Exit Sub
        End If
        End If
        fcmbF03Tro.Enabled = True
        lRStUsuari.Close
End Sub
Private Sub rlsDesabilitarBotao()
        fcmbF03Tro.Enabled = False
End Sub
Private Sub rlsFormatarChaves()
        If (ftxtNumUsu = "") Then
            ftxtNumUsu = 0
        End If

        fintNumUsu = CInt(rgfSemEdicao(ftxtNumUsu))
End Sub
Private Sub rlsLimparCampos()
        ftxtSenhas = ""
        fcmbF03Tro.Enabled = False
End Sub
