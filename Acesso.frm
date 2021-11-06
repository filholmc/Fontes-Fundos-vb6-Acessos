VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form formAcesso 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   ClientHeight    =   3555
   ClientLeft      =   3480
   ClientTop       =   1755
   ClientWidth     =   7290
   ControlBox      =   0   'False
   Icon            =   "Acesso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox ftxtSenhas 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   5370
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2070
      Width           =   1125
   End
   Begin VB.TextBox ftxtNumUsu 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2460
      MaxLength       =   5
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2070
      Width           =   555
   End
   Begin VB.CommandButton fcmbF12Hom 
      Caption         =   "F12"
      Height          =   255
      Left            =   1740
      TabIndex        =   6
      Top             =   2580
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF11Tab 
      Caption         =   "F11"
      Height          =   255
      Left            =   1260
      TabIndex        =   5
      Top             =   2580
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF09LCA 
      Caption         =   "F9"
      Height          =   255
      Left            =   780
      TabIndex        =   4
      Top             =   2580
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbEscape 
      Caption         =   "Fechar (Esc)"
      Height          =   315
      Left            =   5400
      TabIndex        =   3
      Top             =   3090
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF03Ace 
      Caption         =   "Acessar (F3)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   315
      Left            =   780
      TabIndex        =   2
      Top             =   3090
      Width           =   1120
   End
   Begin ComctlLib.ImageList fimlLogAdm 
      Left            =   780
      Top             =   330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   125
      ImageHeight     =   41
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Acesso.frx":000C
            Key             =   "BANESE"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Acesso.frx":18DE
            Key             =   "BESC"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Acesso.frx":2B48
            Key             =   "BB"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Acesso.frx":3CBA
            Key             =   "BNB"
         EndProperty
      EndProperty
   End
   Begin VB.Line flneLinInf 
      BorderColor     =   &H0000C000&
      BorderWidth     =   4
      Index           =   1
      X1              =   750
      X2              =   6525
      Y1              =   2895
      Y2              =   2880
   End
   Begin VB.Image fimgLogAdm 
      Height          =   615
      Left            =   750
      Top             =   300
      Width           =   1290
   End
   Begin VB.Label flblModulo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Acessos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   525
      Left            =   2610
      LinkTimeout     =   0
      TabIndex        =   10
      Top             =   960
      Width           =   2085
   End
   Begin VB.Line flneLinSup 
      BorderColor     =   &H0000C000&
      BorderWidth     =   4
      Index           =   0
      X1              =   750
      X2              =   6525
      Y1              =   975
      Y2              =   960
   End
   Begin VB.Label flblFundos 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fundos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   780
      Left            =   3960
      LinkTimeout     =   0
      TabIndex        =   9
      Top             =   210
      Width           =   2580
   End
   Begin VB.Label flblSenhas 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
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
      Left            =   4710
      LinkTimeout     =   0
      TabIndex        =   8
      Top             =   2130
      Width           =   555
   End
   Begin VB.Label flblNumero 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Número do Usuário"
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
      Left            =   780
      LinkTimeout     =   0
      TabIndex        =   7
      Top             =   2130
      Width           =   1635
   End
End
Attribute VB_Name = "formAcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private flngPrzExp As Long
Private Sub fcmbEscape_Click()
        gbooCancel = True
        Unload Me
End Sub
Private Sub fcmbF03Ace_Click()
        If (ftxtSenhas = "") Then
            rgfMsgBox "Preencha o campo 'Senha'", MsgErr
            ftxtSenhas.SetFocus
            Exit Sub
        End If

        If (gstrSenhas <> rgfSenhaCp(ftxtSenhas)) Then
            rgfMsgBox "Senha não confere", MsgErr
            ftxtSenhas.SetFocus
            Exit Sub
        End If

            gbytModAtv = 1
            gclsLogado.Incluir gstrNomCmp, gintUsuLog, gintAgeLog, " ", "Acessos", gbytModAtv
            gclsDiario.Incluir gbooForLog, gintUsuLog, gstrNomCmp, _
                                                                     1, gintForAtu, "Acessou", 0, gintUsuLog, " "

        If (gDBCFundos.Errors.Count > 0) Then rgsTratarErro Err, gDBCFundos.Errors, Me

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
        rgsTratarFuncoes KeyCode, Me
End Sub
Private Sub Form_Load()
        gbooCancel = False

        Select Case gintCodAdm
               Case 4
                    fimgLogAdm.Top = 270
                    formAcesso.BackColor = &HE0E0E0
                    flneLinSup.Item(0).BorderColor = &H800000
                    flneLinInf.Item(1).BorderColor = &H800000
                    fimgLogAdm.Picture = fimlLogAdm.ListImages(4).Picture
               Case 7
                    fimgLogAdm.Top = 60
                    flblFundos.ForeColor = &H800000
                    flblModulo.ForeColor = &H800000
                    formAcesso.BackColor = &H80FFFF
                    flneLinSup.Item(0).BorderColor = &HFFFF&
                    flneLinInf.Item(1).BorderColor = &HFFFF&
                    fimgLogAdm.Picture = fimlLogAdm.ListImages(3).Picture
               Case 27
                    flneLinSup.Item(0).BorderColor = &HC0&
                    flneLinInf.Item(1).BorderColor = &HC0&
                    fimgLogAdm.Picture = fimlLogAdm.ListImages(2).Picture
               Case 47
                    fimgLogAdm.Picture = fimlLogAdm.ListImages(1).Picture
        End Select

        rgsCentralizarFormIndependente Me
        rgsPosicionarAjuda Me, gintForAtu, gbooForLog
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
End Sub
Private Sub ftxtSenhas_GotFocus()
        rlsConsultar
End Sub
Private Sub rlsConsultar()
        Dim lRStUsuari As Recordset

        Set lRStUsuari = gclsUsuari.Consultar(gintUsuLog)

        If (lRStUsuari.EOF) Then
            rgfMsgBox "Usuário não cadastrado", MsgErr
            ftxtNumUsu.SetFocus
            rlsLimparCampos
            Exit Sub
        Else
            gstrSenhas = lRStUsuari!Senhas
'           gbooUsuLog = IIf( _
'                        lRStUsuari!TemLog, 1, 0)
            gstrNomUsu = lRStUsuari!NomUsu
            gintAgeLog = lRStUsuari!CodAge

        If (lRStUsuari!Status) Then
            rgfMsgBox "Usuário com Acesso Bloqueado", MsgErr
            ftxtNumUsu.SetFocus
            Exit Sub
        End If

            gdatServBD = rgfDataDoServidor
            flngPrzExp = DateDiff("d", gdatServBD, lRStUsuari!DatVal)

        If (gdatServBD > lRStUsuari!DatVal) Then
            rgfMsgBox "Usuário com Acesso Expirado", MsgErr
            ftxtNumUsu.SetFocus
            Exit Sub
        Else
        If (flngPrzExp < 40) Then rgfMsgBox "Acesso deste Usuário expira em " & flngPrzExp & " dia(s)", MsgInf
        End If

        If (gclsAceMod.Ausente(gintUsuLog, 1)) Then
            rgfMsgBox "Usuário não possui Acesso a este Módulo", MsgErr
            ftxtNumUsu.SetFocus
            Exit Sub
        End If
        End If
        fcmbF03Ace.Enabled = True
        lRStUsuari.Close
End Sub
Private Sub rlsDesabilitarBotao()
        fcmbF03Ace.Enabled = False
End Sub
Private Sub rlsFormatarChaves()
        If (ftxtNumUsu = "") Then
            ftxtNumUsu = 0
        End If

        gintUsuLog = CInt(rgfSemEdicao(ftxtNumUsu))
End Sub
Private Sub rlsLimparCampos()
        ftxtSenhas = ""
        fcmbF03Ace.Enabled = False
End Sub
