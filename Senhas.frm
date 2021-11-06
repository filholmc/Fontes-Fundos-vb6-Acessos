VERSION 5.00
Begin VB.Form formSenhas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Troca de Senha"
   ClientHeight    =   3210
   ClientLeft      =   4260
   ClientTop       =   1755
   ClientWidth     =   5670
   Icon            =   "Senhas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Senhas.frx":030A
   ScaleHeight     =   3210
   ScaleWidth      =   5670
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   0
      TabIndex        =   14
      Top             =   2640
      Width           =   5655
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   510
      TabIndex        =   13
      Top             =   450
      Width           =   5145
   End
   Begin VB.Frame frmeFrame1 
      Caption         =   "Informe a Senha Atual:"
      Height          =   915
      Left            =   570
      TabIndex        =   0
      Top             =   600
      Width           =   4515
      Begin VB.TextBox ftxtSenAtu 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   690
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label flblSenAtu 
         AutoSize        =   -1  'True
         Caption         =   "Senha"
         Height          =   195
         Left            =   120
         LinkTimeout     =   0
         TabIndex        =   12
         Top             =   420
         Width           =   465
      End
   End
   Begin VB.Frame frmeFrame2 
      Caption         =   "Informe a Nova Senha duas vezes:"
      Height          =   915
      Left            =   570
      TabIndex        =   2
      Top             =   1620
      Width           =   4515
      Begin VB.TextBox ftxtSenha2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2970
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox ftxtSenha1 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   690
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Senha"
         Height          =   195
         Left            =   2400
         LinkTimeout     =   0
         TabIndex        =   11
         Top             =   420
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Senha"
         Height          =   195
         Left            =   120
         LinkTimeout     =   0
         TabIndex        =   10
         Top             =   420
         Width           =   465
      End
   End
   Begin VB.CommandButton fcmbF09LCA 
      Caption         =   "F9"
      Height          =   255
      Left            =   4230
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF11Tab 
      Caption         =   "F11"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF12Hom 
      Caption         =   "F12"
      Height          =   255
      Left            =   5130
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbEscape 
      Caption         =   "Fechar (Esc)"
      Height          =   315
      Left            =   4440
      TabIndex        =   6
      Top             =   2790
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF03Tro 
      Caption         =   "Trocar (F3)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   2790
      Width           =   1120
   End
End
Attribute VB_Name = "formSenhas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private fbooForLog As Boolean

Private fintForAtu As Integer
Private Function rlfValidacaoDeCamposOk() As Boolean
        rlfValidacaoDeCamposOk = False

        If (Len(ftxtSenha1) < 4) Then
            rgfMsgBox "A Nova Senha deve conter pelo menos 4 caracteres", MsgErr, Me.HelpContextID
            ftxtSenha1.SetFocus
            Exit Function
        End If

        If (ftxtSenha1 = ftxtSenAtu) Then
            rgfMsgBox "Nova Senha é igual à Atual", MsgErr
            ftxtSenha1.SetFocus
            Exit Function
        End If

        If (ftxtSenha1 <> ftxtSenha2) Then
            rgfMsgBox "Novas Senhas não estão iguais", MsgErr
            ftxtSenha1.SetFocus
            Exit Function
        End If

        rlfValidacaoDeCamposOk = True
End Function
Private Sub fcmbEscape_Click()
        Unload Me
End Sub
Private Sub fcmbF03Tro_Click()
        If (Not rlfValidacaoDeCamposOk) Then Exit Sub

            gstrSenhas = rgfSenhaCp(ftxtSenha1)

            gclsUsuari.AlterarSenha gintUsuLog, gstrSenhas

            gclsDiario.Incluir _
                                    fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                               "Trocou", 0, gintUsuLog, " "
        If (gDBCFundos.Errors.Count > 0) Then
            rgsTratarErro Err, gDBCFundos.Errors, Me
        Else
            rgfMsgBox "Senha trocada", MsgInf
        End If
End Sub
Private Sub fcmbF09LCA_Click()
        If (Not TypeOf ActiveControl Is CommandButton) Then ActiveControl.Text = ""
End Sub
Private Sub fcmbF11Tab_Click()
        SendKeys "+{TAB}"
End Sub
Private Sub fcmbF12Hom_Click()
        ftxtSenAtu.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        gintTempoP = 0
        rgsTratarFuncoes KeyCode, Me
End Sub
Private Sub Form_Load()
        gintTempoP = 0
        formMDIAce.fsbrModAce.Panels(4).Picture = IIf(gbooUsuLog Or fbooForLog, _
                                                      formMDIAce.fimlStaBar.ListImages(1).Picture, LoadPicture())
        rgsCentralizarFormIndependente Me
        rgsPosicionarAjuda Me, fintForAtu, fbooForLog

        If (gstrSenhas = "+") Then rgfMsgBox "A sua Senha está igual à Senha Inicial (*)." & vbCr & vbCr & _
                                             "Isto se deve a um dos Motivos abaixo:" & vbCr & _
                                             "- É o seu Primeiro Acesso ou" & vbCr & _
                                             "- A sua Senha foi Restaurada." & vbCr & vbCr & _
                                             "É necessário a Troca de Senha para prosseguir operando.", MsgInf
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        gintTempoP = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
        formMDIAce.fsbrModAce.Panels(4).Picture = LoadPicture()
End Sub
Private Sub ftxtSenUsu_GotFocus()
        fcmbF03Tro.Enabled = False
End Sub
Private Sub ftxtSenAtu_LostFocus()
        If (gstrSenhas <> rgfSenhaCp(ftxtSenAtu)) Then
            rgfMsgBox "Senha Atual não confere", MsgErr
            ftxtSenAtu.SetFocus
        Else
            fcmbF03Tro.Enabled = True
        End If
End Sub
Private Sub rlsLimparCampos()
        ftxtSenAtu = ""
        ftxtSenha1 = ""
        ftxtSenha2 = ""
End Sub
