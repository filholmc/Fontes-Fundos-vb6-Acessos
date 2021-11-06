VERSION 5.00
Begin VB.Form formSenhaR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Restauração de Senha"
   ClientHeight    =   3000
   ClientLeft      =   4530
   ClientTop       =   1755
   ClientWidth     =   5130
   Icon            =   "SenhaR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "SenhaR.frx":030A
   ScaleHeight     =   3000
   ScaleWidth      =   5130
   Begin VB.TextBox ftxtNumUsu 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3030
      MaxLength       =   5
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1035
      Width           =   555
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   0
      TabIndex        =   7
      Top             =   2430
      Width           =   5115
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   510
      TabIndex        =   6
      Top             =   450
      Width           =   4605
   End
   Begin VB.CommandButton fcmbF09LCA 
      Caption         =   "F9"
      Height          =   255
      Left            =   3690
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF11Tab 
      Caption         =   "F11"
      Height          =   255
      Left            =   4140
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF12Hom 
      Caption         =   "F12"
      Height          =   255
      Left            =   4590
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbEscape 
      Caption         =   "Fechar (Esc)"
      Height          =   315
      Left            =   3900
      TabIndex        =   2
      Top             =   2580
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF03Res 
      Caption         =   "Restaurar(F3)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   2580
      Width           =   1120
   End
   Begin VB.Label flblNomUsu 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2490
      TabIndex        =   9
      ToolTipText     =   "Nome do Usuário"
      Top             =   1815
      Width           =   135
   End
   Begin VB.Label flblNumero 
      AutoSize        =   -1  'True
      Caption         =   "Número do Usuário"
      Height          =   195
      Left            =   1530
      LinkTimeout     =   0
      TabIndex        =   8
      Top             =   1095
      Width           =   1365
   End
End
Attribute VB_Name = "formSenhaR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private fbooForLog As Boolean

Private fintForAtu As Integer

Private fintNumUsu As Integer
Private Sub fcmbEscape_Click()
        Unload Me
End Sub
Private Sub fcmbF03Res_Click()
            gclsUsuari.AlterarSenha fintNumUsu, "+"

            gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                      "Restaurou", 0, Format(fintNumUsu, "0"), " "
        If (gDBCFundos.Errors.Count > 0) Then
            rgsTratarErro Err, gDBCFundos.Errors, Me
        Else
            rgfMsgBox "Senha Restaurada", MsgInf
            ftxtNumUsu.SetFocus
        End If
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
Private Sub Form_Activate()
        gintTempoP = 0
        formMDIAce.fsbrModAce.Panels(4).Picture = IIf(gbooUsuLog Or fbooForLog, _
                                                      formMDIAce.fimlStaBar.ListImages(1).Picture, LoadPicture())
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        gintTempoP = 0
        rgsTratarFuncoes KeyCode, Me
End Sub
Private Sub Form_Load()
        rgsCentralizarForm Me
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
        rlsConsultar
End Sub
Private Sub rlsConsultar()
        Dim lRStUsuari As Recordset

        Set lRStUsuari = gclsUsuari.Consultar(fintNumUsu)

        If (lRStUsuari.EOF) Then
            rlsLimparCampos
        Else
            flblNomUsu = lRStUsuari!NomUsu
            fcmbF03Res.Enabled = True
            fcmbF03Res.SetFocus
        End If
        lRStUsuari.Close
End Sub
Private Sub rlsDesabilitarBotao()
        fcmbF03Res.Enabled = False
End Sub
Private Sub rlsFormatarChaves()
        If (ftxtNumUsu = "") Then
            ftxtNumUsu = 0
        End If

        fintNumUsu = CInt(rgfSemEdicao(ftxtNumUsu))
End Sub
Private Sub rlsLimparCampos()
        flblNomUsu = ""
End Sub
