VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form formTempor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Temporização da Proteção"
   ClientHeight    =   5790
   ClientLeft      =   2775
   ClientTop       =   1755
   ClientWidth     =   9420
   Icon            =   "Tempor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   9420
   Begin VB.CheckBox fchkSemPro 
      Caption         =   "Não ativar a Proteção de Acesso"
      Height          =   195
      Left            =   3570
      TabIndex        =   2
      Top             =   2850
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   240
      TabIndex        =   9
      Top             =   5220
      Width           =   8955
   End
   Begin MSComCtl2.UpDown fupdQtdMin 
      Height          =   315
      Left            =   4425
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1620
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Value           =   10
      AutoBuddy       =   -1  'True
      BuddyControl    =   "ftxtQtdMin"
      BuddyDispid     =   196612
      OrigLeft        =   3480
      OrigTop         =   420
      OrigRight       =   3720
      OrigBottom      =   705
      Increment       =   2
      Max             =   30
      Min             =   2
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox ftxtTxtAlt 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   3570
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Tempor.frx":030A
      Top             =   360
      Width           =   5505
   End
   Begin VB.TextBox ftxtQtdMin 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4110
      Locked          =   -1  'True
      MaxLength       =   2
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1620
      Width           =   315
   End
   Begin VB.CommandButton fcmbF11Tab 
      Caption         =   "F11"
      Height          =   255
      Left            =   3570
      TabIndex        =   5
      Top             =   4860
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbF03Alt 
      Caption         =   "Alterar (F3)"
      Default         =   -1  'True
      Height          =   315
      Left            =   6840
      TabIndex        =   3
      Top             =   5370
      Width           =   1120
   End
   Begin VB.CommandButton fcmbF12Hom 
      Caption         =   "F12"
      Height          =   255
      Left            =   4020
      TabIndex        =   6
      Top             =   4860
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton fcmbEscape 
      Caption         =   "Fechar (Esc)"
      Height          =   315
      Left            =   7965
      TabIndex        =   4
      Top             =   5370
      Width           =   1120
   End
   Begin VB.Label flblTempor 
      AutoSize        =   -1  'True
      Caption         =   "Tempo"
      Height          =   195
      Left            =   3570
      LinkTimeout     =   0
      TabIndex        =   8
      Top             =   1680
      Width           =   495
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4815
      Left            =   300
      Picture         =   "Tempor.frx":0383
      Stretch         =   -1  'True
      Top             =   300
      Width           =   3120
   End
End
Attribute VB_Name = "formTempor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private fbooForLog As Boolean

Private fbytQtdMin As Byte

Private fclsParGer As clssParGer

Private fintForAtu As Integer
Private Sub fchkSemPro_Click()
        rlsHabilitarCampos 1 - fchkSemPro
End Sub
Private Sub fcmbEscape_Click()
        Unload Me
End Sub
Private Sub fcmbF03Alt_Click()
        gDBCFundos.BeginTrans
                   fclsParGer.Alterar "TempoDeEspera", ftxtQtdMin

                   gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                               "Alterou", 0, "Tempo", ftxtQtdMin
               If (gDBCFundos.Errors.Count > 0) Then GoTo Erro_DB

                   fclsParGer.Alterar "ProtegeAcesso", IIf(fchkSemPro, "Não", "Sim")

                   gclsDiario.Incluir fbooForLog, gintUsuLog, gstrNomCmp, 1, fintForAtu, _
                                                               "Alterou", 0, "Proteção", IIf(fchkSemPro, "Não", "Sim")
               If (gDBCFundos.Errors.Count > 0) Then GoTo Erro_DB
        gDBCFundos.CommitTrans

        rgfMsgBox "Parâmetros alterados", MsgInf

        gintTempoE = CInt(ftxtQtdMin) * 60

        fcmbF12Hom_Click
        Exit Sub

Erro_DB:
        gDBCFundos.RollbackTrans

        rgsTratarErro Err, gDBCFundos.Errors, Me
End Sub
Private Sub fcmbF11Tab_Click()
        SendKeys "+{TAB}"
End Sub
Private Sub fcmbF12Hom_Click()
        If (fchkSemPro = 1) Then
            fchkSemPro.SetFocus
        Else
            ftxtQtdMin.SetFocus
        End If
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

        Set fclsParGer = New clssParGer

        rlsConsultar
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        gintTempoP = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
        Set fclsParGer = Nothing
        formMDIAce.fsbrModAce.Panels(4).Picture = LoadPicture()
End Sub
Private Sub ftxtQtdMin_GotFocus()
        ftxtQtdMin.SelStart = Len(ftxtQtdMin)
End Sub
Private Sub ftxtQtdMin_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
               Case 38
                If (fbytQtdMin < fupdQtdMin.Max) Then
                    fbytQtdMin = fbytQtdMin + 2
                End If
               Case 40
                If (fbytQtdMin > fupdQtdMin.Min) Then
                    fbytQtdMin = fbytQtdMin - 2
                End If
        End Select
        ftxtQtdMin = fbytQtdMin

        ftxtQtdMin.SelStart = Len(ftxtQtdMin)
End Sub
Private Sub rlsConsultar()
        Dim lRStParGer As Recordset

        Set lRStParGer = fclsParGer.Consultar("TempoDeEspera")

            ftxtQtdMin = lRStParGer!Cteudo
            fbytQtdMin = lRStParGer!Cteudo

        Set lRStParGer = fclsParGer.Consultar("ProtegeAcesso")

            fchkSemPro = IIf(lRStParGer!Cteudo = "Sim", 0, 1)

        fchkSemPro_Click
        lRStParGer.Close
End Sub
Private Sub rlsHabilitarCampos(ByVal vbooStatus As Boolean)
        gbooProAce = vbooStatus
        ftxtTxtAlt.Enabled = vbooStatus
        flblTempor.Enabled = vbooStatus
        ftxtQtdMin.Enabled = vbooStatus
        fupdQtdMin.Enabled = vbooStatus
End Sub
