VERSION 5.00
Begin VB.Form formProtge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proteção de Acesso"
   ClientHeight    =   5790
   ClientLeft      =   2790
   ClientTop       =   1755
   ClientWidth     =   9420
   ControlBox      =   0   'False
   Icon            =   "Protge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   9420
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   240
      TabIndex        =   8
      Top             =   5220
      Width           =   8955
   End
   Begin VB.TextBox ftxtSenhas 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   4140
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2490
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   3570
      MultiLine       =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "Protge.frx":000C
      Top             =   360
      Width           =   5505
   End
   Begin VB.CommandButton fcmbF09LCA 
      Caption         =   "F9"
      Height          =   255
      Left            =   3570
      TabIndex        =   3
      Top             =   4860
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF11Tab 
      Caption         =   "F11"
      Height          =   255
      Left            =   4020
      TabIndex        =   4
      Top             =   4860
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF03Ace 
      Caption         =   "Acessar (F3)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   315
      Left            =   6840
      TabIndex        =   1
      Top             =   5370
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF12Hom 
      Caption         =   "F12"
      Height          =   255
      Left            =   4470
      TabIndex        =   5
      Top             =   4860
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbEscape 
      Caption         =   "Fechar (Esc)"
      Height          =   315
      Left            =   7965
      TabIndex        =   2
      Top             =   5370
      Width           =   1120
   End
   Begin VB.Label flblSenhas 
      AutoSize        =   -1  'True
      Caption         =   "Senha"
      Height          =   195
      Left            =   3570
      LinkTimeout     =   0
      TabIndex        =   7
      Top             =   2550
      Width           =   465
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4815
      Left            =   300
      Picture         =   "Protge.frx":00F2
      Stretch         =   -1  'True
      Top             =   300
      Width           =   3120
   End
End
Attribute VB_Name = "formProtge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Private Sub fcmbEscape_Click()
        If (rgfMsgBox("Confirma o encerramento desta Sessão do Sistema?", MsgNao) = vbYes) Then
            gbooCancel = True
            Unload Me
        End If
End Sub
Private Sub fcmbF03Ace_Click()
        If (ftxtSenhas = "") Then
            rgfMsgBox "Preencha o campo 'Senha'", MsgErr, Me.HelpContextID
            ftxtSenhas.SetFocus
            Exit Sub
        End If

        If (gstrSenhas <> rgfSenhaCp(ftxtSenhas)) Then
            rgfMsgBox "Senha não confere", MsgErr
            ftxtSenhas.SetFocus
            Exit Sub
        End If

        Unload Me
End Sub
Private Sub fcmbF09LCA_Click()
        If (Not TypeOf ActiveControl Is CommandButton) Then ActiveControl.Text = ""
End Sub
Private Sub fcmbF11Tab_Click()
        SendKeys "+{TAB}"
End Sub
Private Sub fcmbF12Hom_Click()
        ftxtSenhas.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        rgsTratarFuncoes KeyCode, Me
End Sub
Private Sub Form_Load()
        gintTempoP = 0
        rgsCentralizarFormIndependente Me
        rgsPosicionarAjuda Me, gintForAtu, gbooForLog
        formMDIAce.fsbrModAce.Panels(4).Picture = LoadPicture()
End Sub
Private Sub Form_Unload(Cancel As Integer)
        gintTempoP = 0
End Sub
Private Sub ftxtSenhas_Change()
        fcmbF03Ace.Enabled = True
End Sub
